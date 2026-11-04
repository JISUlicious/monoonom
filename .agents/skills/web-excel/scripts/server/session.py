"""Server-authoritative document sessions (P1).

The server owns each open workbook. Browsers and the agent CLI send *ops*; the server applies them
in arrival order, assigns a monotonically increasing version, logs them and broadcasts them over
SSE. Conflict rule: last writer wins per cell / per sheet property (the server orders events, so
no timestamps are needed — cf. Figma's multiplayer design). Batches that move coordinates
(row/column insert/delete, sort, sheet structure) are rejected with 409 when another client
changed the document since the batch's base version; that client then resyncs.
"""
from __future__ import annotations

import copy
import json
import os
import queue
import secrets
import threading
import time
from collections import deque

from . import fileio

META_KEYS = ("name", "origName", "state", "tabColor", "showGrid", "defaultColW", "defaultRowH",
             "cols", "rows", "filterHidden", "merges", "freeze", "autoFilter", "charts")
LOG_LIMIT = 5000


class SessionError(Exception):
    def __init__(self, status, message, **extra):
        super().__init__(message)
        self.status, self.message, self.extra = status, message, extra


def new_id(prefix):
    return prefix + secrets.token_hex(5)


def _file_sig(path):
    try:
        st = os.stat(path)
        return [st.st_mtime_ns, st.st_size]
    except OSError:
        return None


class Hub:
    """Fan-out of SSE events to subscribers (each subscriber is a Queue)."""

    def __init__(self):
        self.subs = set()
        self.lock = threading.Lock()

    def subscribe(self, q=None):
        q = q or queue.Queue(maxsize=10000)
        with self.lock:
            self.subs.add(q)
        return q

    def unsubscribe(self, q):
        with self.lock:
            self.subs.discard(q)

    def publish(self, event, data):
        msg = (event, data)
        with self.lock:
            dead = []
            for q in self.subs:
                try:
                    q.put_nowait(msg)
                except queue.Full:
                    dead.append(q)
            for q in dead:  # a stuck client gets dropped and will resync on reconnect
                self.subs.discard(q)

    def __len__(self):
        return len(self.subs)


class DocSession:
    def __init__(self, model: dict, path=None, fmt="xlsx"):
        self.id = new_id("d")
        self.lock = threading.RLock()
        self.version = 0
        self.log: deque = deque(maxlen=LOG_LIMIT)
        self.hub = Hub()
        self.presence: dict = {}
        self.calc: dict = {}            # (sheetId, r, c) -> value reported by a calc source
        self.calc_version = -1
        self.calc_source = None
        self.dirty = False
        self.path = path
        self.base_path = path
        self.format = fmt
        self.file_sig = _file_sig(path) if path else None
        self.name = model.get("name") or (os.path.basename(path) if path else "통합 문서.xlsx")
        self.default_font = model.get("defaultFont") or {"name": "맑은 고딕", "size": 11}
        self.defined_names = model.get("definedNames") or {}
        self.active_sheet = model.get("activeSheet", 0)
        self.warnings = model.get("warnings", [])
        styles = model.get("styles", [])
        self.sheets = [self._sheet_from_json(sj, styles) for sj in model.get("sheets", [])]
        if not self.sheets:
            self.sheets = [self._sheet_from_json({"name": "Sheet1"}, [])]

    # ------------------------------------------------------------ conversion
    @staticmethod
    def _sheet_from_json(sj: dict, styles: list) -> dict:
        sh = {k: copy.deepcopy(sj.get(k)) for k in META_KEYS}
        sh["id"] = sj.get("id") or new_id("s")
        sh["name"] = sj.get("name") or "Sheet1"
        sh["state"] = sh["state"] or "visible"
        sh["showGrid"] = sj.get("showGrid", True) is not False
        sh["defaultColW"] = sh["defaultColW"] or 64
        sh["defaultRowH"] = sh["defaultRowH"] or 20
        sh["cols"] = {str(k): v for k, v in (sh["cols"] or {}).items()}
        sh["rows"] = {str(k): v for k, v in (sh["rows"] or {}).items()}
        sh["merges"] = sh["merges"] or []
        sh["charts"] = sh["charts"] or []
        for ch in sh["charts"]:
            ch.setdefault("id", new_id("ch"))
        fh = sh["filterHidden"]
        if fh is None:  # from a file: rows hidden inside a filtered range belong to the filter
            fh = []
            af = sh.get("autoFilter")
            if af and af.get("cols") and af.get("ref"):
                r1, r2 = _ref_rows(af["ref"])
                for key, e in list(sh["rows"].items()):
                    r = int(key)
                    if e.get("hidden") and r1 < r <= r2:
                        fh.append(r)
                        e.pop("hidden", None)
                        if not e:
                            del sh["rows"][key]
        sh["filterHidden"] = sorted(set(fh))
        cells = {}
        for r, c, it in sj.get("cells", []) or []:
            cell = dict(it)
            s = cell.get("s")
            if isinstance(s, int):
                if 0 <= s < len(styles) and styles[s]:
                    cell["s"] = styles[s]
                else:
                    cell.pop("s", None)
            cell.pop("cv", None)
            cells[(int(r), int(c))] = cell
        sh["cells"] = cells
        return sh

    def sheet_json(self, sh, *, with_cells=True) -> dict:
        out = {k: copy.deepcopy(sh.get(k)) for k in META_KEYS}
        out["id"] = sh["id"]
        if with_cells:
            out["cells"] = [[r, c, cell] for (r, c), cell in sorted(sh["cells"].items())]
        return out

    def snapshot(self) -> dict:
        with self.lock:
            return {
                "docId": self.id, "version": self.version, "name": self.name, "path": self.path,
                "format": self.format, "dirty": self.dirty, "defaultFont": self.default_font,
                "definedNames": self.defined_names, "activeSheet": self.active_sheet,
                "warnings": self.warnings, "styles": [],
                "sheets": [self.sheet_json(sh) for sh in self.sheets],
            }

    def summary(self) -> dict:
        with self.lock:
            return {"docId": self.id, "name": self.name, "path": self.path, "format": self.format,
                    "version": self.version, "dirty": self.dirty, "clients": len(self.hub),
                    "sheets": [sh["name"] for sh in self.sheets],
                    "calcVersion": self.calc_version}

    def file_model(self) -> dict:
        """Model in the shape server.xlsx_io / xls_io expect (style table, hidden rows merged)."""
        with self.lock:
            styles, index = [], {}
            sheets = []
            for sh in self.sheets:
                rows = copy.deepcopy(sh["rows"])
                for r in sh["filterHidden"]:
                    rows.setdefault(str(r), {})["hidden"] = True
                cells = []
                for (r, c), cell in sorted(sh["cells"].items()):
                    it = {k: v for k, v in cell.items() if k != "s"}
                    st = cell.get("s")
                    if st:
                        key = json.dumps(st, sort_keys=True)
                        if key not in index:
                            index[key] = len(styles)
                            styles.append(st)
                        it["s"] = index[key]
                    if cell.get("f"):
                        cv = self.calc.get((sh["id"], r, c), cell.get("v"))
                        if isinstance(cv, dict) and "err" in cv:
                            cv = cv["err"]
                        if cv is not None:
                            it["cv"] = cv
                        it.pop("v", None)
                    cells.append([r, c, it])
                charts = []
                for ch in sh["charts"]:
                    ch2 = {k: v for k, v in ch.items() if k not in ("id", "_range")}
                    charts.append(ch2)
                sheets.append({**{k: copy.deepcopy(sh.get(k)) for k in META_KEYS if k != "filterHidden"},
                               "rows": rows, "cells": cells, "charts": charts})
            return {"name": self.name, "format": self.format, "activeSheet": self.active_sheet,
                    "styles": styles, "definedNames": self.defined_names,
                    "defaultFont": self.default_font, "sheets": sheets}

    def sheet(self, ref) -> dict:
        """Find a sheet by id or (case-insensitive) name."""
        for sh in self.sheets:
            if sh["id"] == ref:
                return sh
        lc = str(ref).lower()
        for sh in self.sheets:
            if sh["name"].lower() == lc:
                return sh
        raise SessionError(404, f"시트를 찾을 수 없습니다: {ref}")

    # ------------------------------------------------------------ ops
    def apply(self, ops, *, actor="user", client_id=None, batch_id=None, label="",
              base_version=None, structural=False) -> int:
        with self.lock:
            if base_version is not None and base_version < self.version:
                since = [e for e in self.log if e["v"] > base_version and e["clientId"] != client_id]
                if since and (structural or any(e["structural"] for e in since)):
                    raise SessionError(409, "다른 클라이언트가 문서 구조를 바꿨습니다. 최신 상태로 다시 불러오세요.",
                                       currentVersion=self.version, reason="structural-conflict")
            if self.version and self.log and base_version is not None and base_version < self.log[0]["v"] - 1:
                raise SessionError(409, "변경 기록이 너무 오래되었습니다. 다시 불러오세요.",
                                   currentVersion=self.version, reason="log-truncated")
            undo = self._snapshot_for(ops)
            for op in ops:
                self._apply_one(op)
            self.version += 1
            entry = {"v": self.version, "actor": actor, "clientId": client_id,
                     "batchId": batch_id or new_id("b"), "label": label, "ops": ops,
                     "structural": bool(structural), "ts": time.time(), "undo": undo}
            self.log.append(entry)
            self.dirty = True
            self.hub.publish("ops", {"docId": self.id, **{k: entry[k] for k in ("v", "actor", "clientId", "batchId", "label", "ops", "structural")}})
            return self.version

    def _snapshot_for(self, ops):
        """Inverse ops (previous state) so an agent batch can be undone later."""
        inv = []
        for op in ops:
            kind = op.get("op")
            try:
                if kind == "setCells":
                    sh = self.sheet(op["sheet"])
                    inv.append({"op": "setCells", "sheet": sh["id"],
                                "cells": [[r, c, copy.deepcopy(sh["cells"].get((r, c)))] for r, c, _ in op["cells"]]})
                elif kind == "setSheetMeta":
                    sh = self.sheet(op["sheet"])
                    inv.append({"op": "setSheetMeta", "sheet": sh["id"],
                                "meta": {k: copy.deepcopy(sh.get(k)) for k in op["meta"] if k in META_KEYS}})
                elif kind == "sheets":
                    inv.append({"op": "sheets", "order": [sh["id"] for sh in self.sheets],
                                "add": {sh["id"]: self.sheet_json(sh) for sh in self.sheets}})
                elif kind == "setWorkbook":
                    inv.append({"op": "setWorkbook", "props": {"definedNames": copy.deepcopy(self.defined_names),
                                                               "activeSheet": self.active_sheet}})
            except SessionError:
                pass
        return list(reversed(inv))

    def _apply_one(self, op):
        kind = op.get("op")
        if kind == "setCells":
            sh = self.sheet(op["sheet"])
            for r, c, cell in op["cells"]:
                if cell:
                    sh["cells"][(int(r), int(c))] = dict(cell)
                else:
                    sh["cells"].pop((int(r), int(c)), None)
        elif kind == "setSheetMeta":
            sh = self.sheet(op["sheet"])
            for k, v in op["meta"].items():
                if k in META_KEYS:
                    sh[k] = copy.deepcopy(v)
            for ch in sh["charts"] or []:
                ch.setdefault("id", new_id("ch"))
        elif kind == "sheets":
            by_id = {sh["id"]: sh for sh in self.sheets}
            added = op.get("add") or {}
            order = []
            for sid in op["order"]:
                if sid in added and sid not in by_id:
                    sj = dict(added[sid], id=sid)
                    order.append(self._sheet_from_json(sj, []))
                elif sid in by_id:
                    order.append(by_id[sid])
                elif sid in added:
                    order.append(self._sheet_from_json(dict(added[sid], id=sid), []))
                else:
                    raise SessionError(400, f"알 수 없는 시트 id: {sid}")
            if not order:
                raise SessionError(400, "시트가 하나 이상 있어야 합니다.")
            self.sheets = order
            self.active_sheet = min(self.active_sheet, len(order) - 1)
        elif kind == "setWorkbook":
            props = op.get("props", {})
            if "definedNames" in props:
                self.defined_names = props["definedNames"] or {}
            if "activeSheet" in props:
                self.active_sheet = int(props["activeSheet"])
        else:
            raise SessionError(400, f"알 수 없는 op: {kind}")

    def events_since(self, since):
        with self.lock:
            if self.log and since < self.log[0]["v"] - 1:
                return None  # caller must resync
            return [{"docId": self.id, **{k: e[k] for k in ("v", "actor", "clientId", "batchId", "label", "ops", "structural")}}
                    for e in self.log if e["v"] > since]

    def set_calc(self, version, values, source, full=False):
        with self.lock:
            if version < self.calc_version:
                return
            if full:
                self.calc = {}
            for sid, r, c, v in values:
                self.calc[(sid, int(r), int(c))] = v
            self.calc_version = version
            self.calc_source = source


class Registry:
    """All open documents + a global hub (doc opened/closed/saved) for the page's tab bar."""

    def __init__(self):
        self.docs: dict[str, DocSession] = {}
        self.lock = threading.RLock()
        self.hub = Hub()

    def list(self):
        with self.lock:
            return [d.summary() for d in self.docs.values()]

    def get(self, doc_id) -> DocSession:
        d = self.docs.get(doc_id)
        if not d:
            raise SessionError(404, f"열린 문서가 없습니다: {doc_id}")
        return d

    def find_path(self, path):
        with self.lock:
            for d in self.docs.values():
                if d.path and os.path.normcase(d.path) == os.path.normcase(path):
                    return d
        return None

    def add(self, doc: DocSession):
        with self.lock:
            self.docs[doc.id] = doc
        self.hub.publish("opened", doc.summary())
        return doc

    def remove(self, doc_id):
        with self.lock:
            d = self.docs.pop(doc_id, None)
        if d:
            d.hub.publish("closed", {"docId": doc_id})
            self.hub.publish("closed", {"docId": doc_id})

    def dirty_names(self):
        with self.lock:
            return [d.name for d in self.docs.values() if d.dirty]

    def write_recovery(self, state_dir):
        os.makedirs(os.path.join(state_dir, "recovery"), exist_ok=True)
        with self.lock:
            for d in self.docs.values():
                if d.dirty:
                    p = os.path.join(state_dir, "recovery", f"{d.id}.json")
                    with open(p, "w", encoding="utf-8") as fh:
                        json.dump({"path": d.path, "savedAt": time.time(), "workbook": d.file_model()},
                                  fh, ensure_ascii=False, default=str)


def backup_file(path, state_dir):
    """Copy the current file into .web-excel/backups before overwriting (through fileio: DRM-safe)."""
    if not os.path.isfile(path):
        return None
    bdir = os.path.join(state_dir, "backups")
    os.makedirs(bdir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(bdir, f"{base}.{time.strftime('%Y%m%d-%H%M%S')}{ext}")
    fileio.write_bytes(dst, fileio.read_bytes(path))
    return dst


def _ref_rows(ref):
    import re
    nums = [int(n) for n in re.findall(r"(\d+)", ref)]
    if not nums:
        return 0, 0
    return nums[0] - 1, nums[-1] - 1
