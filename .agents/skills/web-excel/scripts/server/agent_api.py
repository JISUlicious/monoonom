"""Agent-facing API (P2): high-level edits translated into state ops, plus compact reads.

All edits are applied as one batch (actor="agent") so the user sees them live in the preview and
can be undone with POST /api/docs/{id}/undo {batchId}.
"""
from __future__ import annotations

import copy
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import time

from . import a1 as A
from .app import SKILL_DIR, ApiError, route
from .docs_api import REGISTRY, _doc, _wrap
from .formula_compat import _STR_OR_SHEET
from .session import new_id

STYLE_KEYS = {"b", "i", "u", "st", "fs", "ff", "fc", "bg", "ha", "va", "wrap", "ind", "rot", "nf",
              "bl", "br", "bt", "bb"}
EXCEL_EPOCH = dt.datetime(1899, 12, 30)
ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


# ------------------------------------------------------------------ helpers

def resolve(doc, ref, sheet=None):
    """'Sheet!A1:B2' | 'A1:B2' (+ optional sheet) -> (sheet dict, (r1,c1,r2,c2))."""
    if not ref:
        raise ApiError(400, "범위(range)가 필요합니다.")
    sname, rng = A.split_sheet(ref)
    sh = doc.sheet(sname or sheet) if (sname or sheet) else doc.sheets[min(doc.active_sheet, len(doc.sheets) - 1)]
    try:
        return sh, A.parse_range(rng)
    except ValueError as e:
        raise ApiError(400, str(e))


def used_extent(sh, pending=None):
    keys = list(sh["cells"]) + [k for k, v in (pending or {}).items() if v]
    if not keys:
        return -1, -1
    return max(r for r, _ in keys), max(c for _, c in keys)


def clamp(sh, g, *, min_rows=0, pending=None):
    """Cut whole-row/column ranges to the used area (including cells written earlier in this batch)."""
    mr, mc = used_extent(sh, pending)
    r1, c1, r2, c2 = g
    return r1, c1, min(r2, max(mr, r1 + min_rows)), min(c2, max(mc, c1))


def to_serial(d: dt.datetime) -> float:
    return (d - EXCEL_EPOCH).total_seconds() / 86400


def parse_value(x, parse=True):
    """Agent value -> (cell fields dict, implied number format | None). Mirrors typing in Excel."""
    if isinstance(x, dict):
        out, nf = {}, x.get("nf")
        if "f" in x:
            out["f"] = x["f"] if str(x["f"]).startswith("=") else "=" + str(x["f"])
        elif "text" in x:
            out["v"] = str(x["text"])
        elif "date" in x:
            out["v"], nf = to_serial(dt.datetime.fromisoformat(str(x["date"]))), nf or "yyyy-mm-dd"
        elif "v" in x:
            out, nf2 = parse_value(x["v"], parse=False)
            nf = nf or nf2
        return out, nf
    if x is None or x == "":
        return {}, None
    if isinstance(x, bool) or isinstance(x, (int, float)):
        return {"v": x}, None
    s = str(x)
    if s.startswith("=") and len(s) > 1:
        return {"f": s}, None
    if s.startswith("'"):
        return {"v": s[1:]}, None
    if not parse:
        return {"v": s}, None
    t = s.strip()
    if t.upper() in ("TRUE", "FALSE"):
        return {"v": t.upper() == "TRUE"}, None
    if t.upper() in ERRORS:
        return {"v": t.upper(), "t": "e"}, None
    if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d+)?", t):
        v = float(t.replace(",", ""))
        return {"v": int(v) if v.is_integer() else v}, "#,##0.00" if "." in t else "#,##0"
    if re.fullmatch(r"[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?", t):
        v = float(t)
        return {"v": int(v) if v.is_integer() and "e" not in t.lower() and "." not in t else v}, None
    m = re.fullmatch(r"([-+]?[\d,]*\.?\d+)\s*%", t)
    if m:
        return {"v": float(m.group(1).replace(",", "")) / 100}, "0.00%" if "." in m.group(1) else "0%"
    m = re.fullmatch(r"([-+]?)([₩$])\s*([\d,]*\.?\d+)", t)
    if m:
        v = float(m.group(3).replace(",", "")) * (-1 if m.group(1) == "-" else 1)
        return {"v": int(v) if v.is_integer() else v}, '"₩"#,##0' if m.group(2) == "₩" else '"$"#,##0.00'
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\.?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?", t)
    if m:
        try:
            d = dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0))
            return {"v": to_serial(d)}, "yyyy-mm-dd h:mm" if m.group(4) else "yyyy-mm-dd"
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", t)
    if m:
        secs = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3) or 0)
        return {"v": secs / 86400}, "h:mm:ss" if m.group(3) else "h:mm"
    return {"v": s}, None


_REF = re.compile(r"(?<![A-Za-z0-9_.$À-￿])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(À-￿])")


def translate(formula: str, dr: int, dc: int) -> str:
    """Shift relative A1 references (copy/fill semantics). Strings and quoted sheet names are kept."""
    if not dr and not dc:
        return formula

    def shift(m):
        cabs, col, rabs, row = m.groups()
        c, r = A.col_index(col), int(row) - 1
        if c >= A.MAX_COLS:
            return m.group(0)
        if not cabs:
            c += dc
        if not rabs:
            r += dr
        if c < 0 or r < 0:
            return "#REF!"
        return f"{cabs}{A.col_name(c)}{rabs}{r + 1}"

    parts = _STR_OR_SHEET.split(formula)
    return "".join(p if i % 2 else _REF.sub(shift, p) for i, p in enumerate(parts))


def merged_style(cur, patch=None, unset=()):
    st = dict(cur or {})
    for k, v in (patch or {}).items():
        if k not in STYLE_KEYS:
            raise ApiError(400, f"알 수 없는 서식 키: {k} (사용 가능: {', '.join(sorted(STYLE_KEYS))})")
        if v is None or v is False:
            st.pop(k, None)
        else:
            st[k] = v
    for k in unset:
        st.pop(k, None)
    return st or None


def cell_with(cur, fields, nf=None):
    cell = {k: v for k, v in (cur or {}).items() if k in ("s", "h", "cm")}
    cell.update(fields)
    if nf and not (cell.get("s") or {}).get("nf"):
        cell["s"] = dict(cell.get("s") or {}, nf=nf)
    return cell if any(k in cell for k in ("v", "f", "s", "h", "cm")) else None


# ------------------------------------------------------------------ calc freshness

def _has_formulas(doc):
    return any(c.get("f") for sh in doc.sheets for c in sh["cells"].values())


def ensure_calc(doc, wait=2.0):
    """Make formula results current: browser report -> Node worker -> stale."""
    if not _has_formulas(doc):
        return {"status": "none"}
    if doc.calc_version >= doc.version:
        return {"status": "fresh", "source": doc.calc_source}
    if len(doc.hub) and wait:
        end = time.time() + wait
        while time.time() < end:
            time.sleep(0.1)
            if doc.calc_version >= doc.version:
                return {"status": "fresh", "source": doc.calc_source}
    node = shutil.which("node")
    if node:
        worker = os.path.join(SKILL_DIR, "scripts", "calc", "worker.mjs")
        version = doc.version
        try:
            res = subprocess.run([node, worker], input=json.dumps(doc.snapshot(), ensure_ascii=False, default=str),
                                 capture_output=True, text=True, timeout=60, encoding="utf-8")
            if res.returncode == 0:
                doc.set_calc(version, json.loads(res.stdout)["values"], "node", full=True)
                return {"status": "fresh", "source": "node"}
            err = res.stderr[-300:]
        except (OSError, subprocess.TimeoutExpired, ValueError) as e:
            err = str(e)
        return {"status": "stale", "source": doc.calc_source, "note": f"Node 계산 실패: {err}"}
    return {"status": "stale", "source": doc.calc_source,
            "note": "수식 결과가 최신이 아닐 수 있습니다. 미리보기 창을 열거나 Node.js 18+를 설치하면 계산됩니다."}


def cell_value(doc, sh, r, c, cell):
    if cell.get("f"):
        v = doc.calc.get((sh["id"], r, c), cell.get("v"))
        v = v["err"] if isinstance(v, dict) and "err" in v else v
    else:
        v = cell.get("v")
    if isinstance(v, float):
        v = float(f"{v:.15g}")  # Excel keeps 15 significant digits
        if v.is_integer() and abs(v) < 1e15:
            v = int(v)
    return v


def display_value(doc, sh, r, c, cell):
    """Readable value for the agent: date-formatted serials become ISO dates/times."""
    v = cell_value(doc, sh, r, c, cell)
    nf = ((cell.get("s") or {}).get("nf") or "")
    clean = re.sub(r'"[^"]*"|\\.', "", nf)
    if isinstance(v, (int, float)) and not isinstance(v, bool) and clean and clean != "General":
        has_date = bool(re.search(r"[yd]", clean, re.I)) or bool(re.search(r"m{3,}", clean))
        has_time = bool(re.search(r"[hs]", clean, re.I))
        if (has_date or has_time) and 0 <= v < 2958466:
            d = EXCEL_EPOCH + dt.timedelta(days=float(v))
            d = d.replace(microsecond=0)
            if has_date and has_time:
                return d.strftime("%Y-%m-%d %H:%M:%S")
            if has_date:
                return d.strftime("%Y-%m-%d")
            secs = round((float(v) % 1) * 86400)
            return f"{secs // 3600}:{secs % 3600 // 60:02d}:{secs % 60:02d}"
    return v


# ------------------------------------------------------------------ edits

def build_ops(doc, edits):
    """Translate agent edits into state ops. Returns (ops, structural, touched_summary)."""
    ops, structural, summary = [], False, []
    cellbuf: dict = {}  # sheet id -> {(r,c): cell}

    def cur(sh, r, c):
        buf = cellbuf.get(sh["id"], {})
        return buf[(r, c)] if (r, c) in buf else sh["cells"].get((r, c))

    def put(sh, r, c, cell):
        cellbuf.setdefault(sh["id"], {})[(r, c)] = cell

    meta: dict = {}  # sheet id -> meta patch

    def meta_of(sh, key):
        m = meta.setdefault(sh["id"], {})
        if key not in m:
            m[key] = copy.deepcopy(sh[key])
        return m[key]

    for e in edits:
        kind = e.get("op")
        sheet = e.get("sheet")
        if kind in ("set", "write"):
            sh, (r1, c1, r2, c2) = resolve(doc, e.get("range"), sheet)
            vals = e.get("values", e.get("value"))
            parse = e.get("parse", True)
            if not isinstance(vals, list):
                vals = [[vals]]
            elif vals and not isinstance(vals[0], list):
                vals = [vals] if (r1 == r2) else [[v] for v in vals]
            h, w = len(vals), max((len(row) for row in vals), default=0)
            fill = h == 1 and w == 1 and (r2 > r1 or c2 > c1)
            if not fill and (r1, c1) == (r2, c2):
                r2, c2 = r1 + h - 1, c1 + w - 1
            if not fill and (r2 - r1 + 1, c2 - c1 + 1) != (h, w):
                raise ApiError(400, f"값 크기({h}x{w})가 범위({A.range_a1(r1, c1, r2, c2)})와 다릅니다.")
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
                raise ApiError(400, "한 번에 쓸 수 있는 셀은 200,000개까지입니다.")
            for i in range(r2 - r1 + 1):
                for j in range(c2 - c1 + 1):
                    raw = vals[0][0] if fill else (vals[i][j] if j < len(vals[i]) else None)
                    fields, nf = parse_value(raw, parse)
                    if fill and "f" in fields:
                        fields["f"] = translate(fields["f"], i, j)
                    put(sh, r1 + i, c1 + j, cell_with(cur(sh, r1 + i, c1 + j), fields, nf))
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 입력")
        elif kind == "clear":
            sh, g = resolve(doc, e.get("range"), sheet)
            r1, c1, r2, c2 = clamp(sh, g, pending=cellbuf.get(sh["id"]))
            what = e.get("what", "contents")
            for (r, c) in set(sh["cells"]) | set(cellbuf.get(sh["id"], {})):
                if r1 <= r <= r2 and c1 <= c <= c2:
                    cell = cur(sh, r, c) or {}
                    if what == "all":
                        put(sh, r, c, None)
                    elif what == "formats":
                        put(sh, r, c, {k: v for k, v in cell.items() if k != "s"} or None)
                    else:
                        keep = {k: v for k, v in cell.items() if k in ("s", "cm")}
                        put(sh, r, c, keep or None)
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 지우기({what})")
        elif kind == "format":
            sh, g = resolve(doc, e.get("range"), sheet)
            r1, c1, r2, c2 = clamp(sh, g, pending=cellbuf.get(sh["id"]))
            patch, unset = dict(e.get("set") or {}), e.get("unset") or []
            border = patch.pop("border", None) or e.get("border")
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
                raise ApiError(400, "서식은 한 번에 200,000셀까지 적용할 수 있습니다.")
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = dict(cur(sh, r, c) or {})
                    st = merged_style(cell.get("s"), patch, unset)
                    if border:
                        st = apply_border(st or {}, border, r == r1, r == r2, c == c1, c == c2)
                    if st:
                        cell["s"] = st
                    else:
                        cell.pop("s", None)
                    put(sh, r, c, cell or None)
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 서식")
        elif kind in ("merge", "unmerge"):
            sh, (r1, c1, r2, c2) = resolve(doc, e.get("range"), sheet)
            merges = meta_of(sh, "merges")
            keep = []
            for mref in merges:
                a, b, c_, d = A.parse_range(mref)
                if not (a <= r2 and r1 <= c_ and b <= c2 and c1 <= d):
                    keep.append(mref)
            if kind == "merge":
                keep.append(A.range_a1(r1, c1, r2, c2))
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        if (r, c) != (r1, c1) and cur(sh, r, c):
                            cell = cur(sh, r, c)
                            put(sh, r, c, {"s": cell["s"]} if cell.get("s") else None)
                if e.get("center", True):
                    cell = dict(cur(sh, r1, c1) or {})
                    cell["s"] = dict(cell.get("s") or {}, ha="center")
                    put(sh, r1, c1, cell)
            meta[sh["id"]]["merges"] = keep
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} {'병합' if kind == 'merge' else '병합 해제'}")
        elif kind in ("colWidth", "rowHeight", "hide", "unhide"):
            sh, (r1, c1, r2, c2) = resolve(doc, e.get("range") or e.get("cols") or e.get("rows"), sheet)
            axis = "cols" if (kind == "colWidth" or (kind in ("hide", "unhide") and r1 == 0 and r2 == A.MAX_ROWS - 1)) else "rows"
            dims = meta_of(sh, axis)
            lo, hi = (c1, c2) if axis == "cols" else (r1, r2)
            if hi - lo > 5000:
                raise ApiError(400, "한 번에 5,000개 행/열까지 바꿀 수 있습니다.")
            for i in range(lo, hi + 1):
                d = dict(dims.get(str(i)) or {})
                if kind == "colWidth":
                    d["w"] = round(float(e["width"]) * 7 + 5)  # Excel character width -> px
                elif kind == "rowHeight":
                    d["h"] = round(float(e["height"]) * 96 / 72)  # points -> px
                elif kind == "hide":
                    d["hidden"] = True
                else:
                    d.pop("hidden", None)
                if d:
                    dims[str(i)] = d
                else:
                    dims.pop(str(i), None)
            summary.append(f"{sh['name']} {axis} {kind}")
        elif kind == "freeze":
            sh = doc.sheet(sheet) if sheet else doc.sheets[doc.active_sheet]
            cellref = e.get("cell")
            meta.setdefault(sh["id"], {})["freeze"] = (A.a1(*A.parse_range(cellref)[:2]) if cellref and cellref not in ("A1", "none") else None)
            summary.append(f"{sh['name']} 틀 고정 {cellref or '해제'}")
        elif kind == "filter":
            sh, g = (resolve(doc, e["range"], sheet) if e.get("range") else (doc.sheet(sheet) if sheet else doc.sheets[doc.active_sheet], None))
            if g:
                r1, c1, r2, c2 = clamp(sh, g, pending=cellbuf.get(sh["id"]))
                meta.setdefault(sh["id"], {})["autoFilter"] = {"ref": A.range_a1(r1, c1, max(r2, r1 + 1), c2), "cols": {}}
            else:
                meta.setdefault(sh["id"], {})["autoFilter"] = None
            meta[sh["id"]]["filterHidden"] = []
            summary.append(f"{sh['name']} 자동 필터 {'켜기' if g else '끄기'}")
        elif kind == "chart":
            sh_data, g = resolve(doc, e.get("data"), sheet)
            target = doc.sheet(e["sheet"]) if e.get("sheet") else sh_data
            chart = chart_from_range(doc, sh_data, g, e, cellbuf.get(sh_data["id"], {}))
            charts = meta_of(target, "charts")
            charts.append(chart)
            summary.append(f"{target['name']} 차트 추가({chart['type']}, id={chart['id']})")
        elif kind == "deleteChart":
            sh = doc.sheet(sheet) if sheet else doc.sheets[doc.active_sheet]
            charts = meta_of(sh, "charts")
            cid = e.get("id")
            idx = e.get("index")
            new = [ch for i, ch in enumerate(charts) if not (ch.get("id") == cid or i == idx)]
            if len(new) == len(charts):
                raise ApiError(404, "차트를 찾을 수 없습니다.")
            meta[sh["id"]]["charts"] = new
            summary.append(f"{sh['name']} 차트 삭제")
        elif kind == "addSheet":
            name = (e.get("name") or "").strip() or _unique_name(doc, "Sheet")
            _check_sheet_name(doc, name)
            sid = new_id("s")
            order = [s["id"] for s in doc.sheets]
            at = int(e.get("at", len(order)))
            order.insert(max(0, min(at, len(order))), sid)
            ops.append({"op": "sheets", "order": order, "add": {sid: {"name": name, "cells": []}}})
            structural = True
            summary.append(f"시트 추가 '{name}'")
        elif kind == "renameSheet":
            sh = doc.sheet(sheet or e.get("from"))
            name = (e.get("name") or "").strip()
            _check_sheet_name(doc, name, exclude=sh)
            old = sh["name"]
            meta.setdefault(sh["id"], {})["name"] = name
            for other in doc.sheets:
                for (r, c), cell in other["cells"].items():
                    if cell.get("f") and old.lower() in cell["f"].lower():
                        nf = A.rename_sheet_in_formula(cell["f"], old, name)
                        if nf != cell["f"]:
                            c2 = dict(cur(other, r, c) or cell)
                            c2["f"] = nf
                            put(other, r, c, c2)
                charts = other["charts"]
                if any(old.lower() in json.dumps(ch, ensure_ascii=False).lower() for ch in charts):
                    new = copy.deepcopy(meta_of(other, "charts"))
                    for ch in new:
                        for s in ch.get("series", []):
                            for k in ("val", "cat", "nameRef"):
                                if s.get(k):
                                    s[k] = A.rename_sheet_in_formula("=" + s[k], old, name)[1:]
                        ch["dirty"] = True
                    meta[other["id"]]["charts"] = new
            names = {k: A.rename_sheet_in_formula(v if v.startswith("=") else "=" + v, old, name).lstrip("=")
                     for k, v in doc.defined_names.items()}
            ops.append({"op": "setWorkbook", "props": {"definedNames": names}})
            structural = True
            summary.append(f"시트 이름 '{old}' → '{name}'")
        elif kind == "deleteSheet":
            sh = doc.sheet(sheet or e.get("name"))
            if len(doc.sheets) <= 1:
                raise ApiError(400, "마지막 시트는 삭제할 수 없습니다.")
            ops.append({"op": "sheets", "order": [s["id"] for s in doc.sheets if s is not sh], "add": {}})
            structural = True
            summary.append(f"시트 삭제 '{sh['name']}'")
        elif kind == "name":
            names = dict(doc.defined_names)
            if e.get("ref"):
                names[e["name"]] = e["ref"].lstrip("=")
            else:
                names.pop(e["name"], None)
            ops.append({"op": "setWorkbook", "props": {"definedNames": names}})
            summary.append(f"이름 정의 {e['name']}")
        else:
            raise ApiError(400, f"알 수 없는 편집 op: {kind}")

    for sid, cells in cellbuf.items():
        ops.append({"op": "setCells", "sheet": sid, "cells": [[r, c, cell] for (r, c), cell in sorted(cells.items())]})
    for sid, m in meta.items():
        ops.append({"op": "setSheetMeta", "sheet": sid, "meta": m})
    # structure ops last so cell/meta ops still find their sheets (deleted sheets are dropped)
    ops.sort(key=lambda o: o["op"] == "sheets")
    return ops, structural, summary


def apply_border(st, spec, top, bottom, left, right):
    if isinstance(spec, str):
        spec = {"kind": spec}
    kind = spec.get("kind", "all")
    side = {"s": spec.get("style", "thin")}
    if spec.get("color") and spec["color"].upper() not in ("#000", "#000000"):
        side["c"] = spec["color"]
    st = dict(st)
    sel = {"all": (True, True, True, True), "outline": (top, bottom, left, right), "top": (top, False, False, False),
           "bottom": (False, bottom, False, False), "left": (False, False, left, False),
           "right": (False, False, False, right), "inside": (not top, not bottom, not left, not right)}
    if kind == "none":
        for k in ("bt", "bb", "bl", "br"):
            st.pop(k, None)
        return st
    if kind not in sel:
        raise ApiError(400, f"테두리 종류: {', '.join(sel)} 또는 none")
    for k, on in zip(("bt", "bb", "bl", "br"), sel[kind]):
        if on:
            st[k] = side
    return st


def _check_sheet_name(doc, name, exclude=None):
    if not name or len(name) > 31 or re.search(r"[\\/?*\[\]:]", name):
        raise ApiError(400, "시트 이름은 1~31자이고 \\ / ? * [ ] : 를 쓸 수 없습니다.")
    if any(s is not exclude and s["name"].lower() == name.lower() for s in doc.sheets):
        raise ApiError(409, f"같은 이름의 시트가 있습니다: {name}")


def _unique_name(doc, base):
    i = 1
    names = {s["name"].lower() for s in doc.sheets}
    while f"{base}{i}".lower() in names:
        i += 1
    return f"{base}{i}"


def chart_from_range(doc, sh, g, e, pending=None):
    """Same heuristics as the UI (first row = series names, first text/date column = categories)."""
    pending = pending or {}
    r1, c1, r2, c2 = clamp(sh, g, pending=pending)
    get = lambda r, c: (pending[(r, c)] if (r, c) in pending else sh["cells"].get((r, c))) or {}
    val = lambda r, c: cell_value(doc, sh, r, c, get(r, c)) if not get(r, c).get("f") else 0
    is_date = lambda r, c: bool(re.search(r"[ymd]", (get(r, c).get("s") or {}).get("nf", "")))
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    header_row = rows > 1 and any(isinstance(val(r1, c), str) for c in range(c1, c2 + 1))
    dr1 = r1 + (1 if header_row else 0)
    header_col = cols > 1 and any(isinstance(val(r, c1), str) or is_date(r, c1) for r in range(dr1, r2 + 1))
    dc1 = c1 + (1 if header_col else 0)
    by_rows = e.get("byRows")
    if by_rows is None:
        by_rows = (c2 - dc1 + 1) > (r2 - dr1 + 1) and e.get("type") != "scatter"
    q = lambda a, b, c, d: f"{A.quote_sheet(sh['name'])}!${A.col_name(b)}${a + 1}:${A.col_name(d)}${c + 1}"
    series = []
    if not by_rows:
        cat = q(dr1, c1, r2, c1) if header_col else None
        for c in range(dc1, c2 + 1):
            series.append({"name": None, "nameRef": q(r1, c, r1, c) if header_row else None, "cat": cat, "val": q(dr1, c, r2, c)})
    else:
        cat = q(r1, dc1, r1, c2) if header_row else None
        for r in range(dr1, r2 + 1):
            series.append({"name": None, "nameRef": q(r, c1, r, c1) if header_col else None, "cat": cat, "val": q(r, dc1, r, c2)})
    ctype = e.get("type", "bar")
    preset = {"barH": ("bar", "bar", "clustered"), "barStacked": ("bar", "col", "stacked")}.get(ctype)
    t, bar_dir, grouping = preset or (ctype, "col", "standard" if ctype in ("line", "area") else "clustered")
    if t not in ("bar", "line", "area", "pie", "doughnut", "scatter", "radar"):
        raise ApiError(400, "차트 종류: bar, barH, barStacked, line, area, pie, doughnut, scatter, radar")
    at = e.get("at")
    ar, ac = (A.parse_range(at)[:2] if at else (r1, min(c2 + 2, A.MAX_COLS - 10)))
    return {"id": new_id("ch"), "type": t, "barDir": bar_dir, "grouping": grouping, "title": e.get("title"),
            "legend": e.get("legend", "r" if (len(series) > 1 or t in ("pie", "doughnut")) else None),
            "xTitle": e.get("xTitle"), "yTitle": e.get("yTitle"), "series": series, "dirty": True,
            "anchor": {"c": ac, "r": ar, "dx": 0, "dy": 0, "w": int(e.get("width", 480)), "h": int(e.get("height", 288))}}


@route("POST", r"/api/docs/([\w-]+)/edit")
@_wrap
def r_edit(h, m, q, body):
    doc = _doc(m)
    edits = body.get("edits") or body.get("ops") or []
    if isinstance(edits, dict):
        edits = [edits]
    with doc.lock:
        ops, structural, summary = build_ops(doc, edits)
        if not ops:
            return {"version": doc.version, "summary": []}
        batch = body.get("batchId") or new_id("b")
        v = doc.apply(ops, actor="agent", client_id=body.get("clientId", "agent-cli"), batch_id=batch,
                      label=body.get("label") or "; ".join(summary)[:120], structural=structural)
    return {"version": v, "batchId": batch, "summary": summary}


# ------------------------------------------------------------------ reads

@route("GET", r"/api/docs/([\w-]+)/range")
@_wrap
def r_range(h, m, q, body):
    doc = _doc(m)
    ref = (q.get("ref") or [""])[0]
    what = (q.get("what") or ["values"])[0]
    max_rows = int((q.get("max") or ["200"])[0])
    calc = ensure_calc(doc) if what != "formulas" else {"status": "skipped"}
    with doc.lock:
        sh, g = resolve(doc, ref or "A1:Z50", (q.get("sheet") or [None])[0])
        r1, c1, r2, c2 = clamp(sh, g)
        truncated = None
        if r2 - r1 + 1 > max_rows:
            truncated = A.range_a1(r1 + max_rows, c1, r2, c2)
            r2 = r1 + max_rows - 1
        if c2 - c1 + 1 > 100:
            c2 = c1 + 99
        rows = []
        for r in range(r1, r2 + 1):
            row = []
            for c in range(c1, c2 + 1):
                cell = sh["cells"].get((r, c)) or {}
                if what == "values":
                    row.append(display_value(doc, sh, r, c, cell))
                elif what == "raw":
                    row.append(cell_value(doc, sh, r, c, cell))
                elif what == "formulas":
                    row.append(cell.get("f") if cell.get("f") else cell.get("v"))
                else:
                    item = {k: v for k, v in cell.items() if k in ("f", "h", "cm", "t")}
                    v = cell_value(doc, sh, r, c, cell)
                    if v is not None:
                        item["v"] = v
                    if cell.get("s"):
                        item["s"] = cell["s"]
                    row.append(item or None)
            rows.append(row)
        return {"sheet": sh["name"], "ref": A.range_a1(r1, c1, r2, c2) if r2 >= r1 and c2 >= c1 else None,
                "rows": rows, "calc": calc, "truncated": truncated, "version": doc.version}


@route("GET", r"/api/docs/([\w-]+)/describe")
@_wrap
def r_describe(h, m, q, body):
    doc = _doc(m)
    calc = ensure_calc(doc)
    out = {"docId": doc.id, "name": doc.name, "path": doc.path, "format": doc.format, "dirty": doc.dirty,
           "version": doc.version, "calc": calc, "definedNames": doc.defined_names, "sheets": []}
    with doc.lock:
        for i, sh in enumerate(doc.sheets):
            mr, mc = used_extent(sh)
            info = {"name": sh["name"], "active": i == doc.active_sheet, "state": sh["state"],
                    "used": A.range_a1(0, 0, mr, mc) if mr >= 0 else None, "cells": len(sh["cells"]),
                    "formulas": sum(1 for c in sh["cells"].values() if c.get("f")),
                    "merges": len(sh["merges"]), "freeze": sh["freeze"],
                    "autoFilter": (sh["autoFilter"] or {}).get("ref"),
                    "charts": [{"id": ch.get("id"), "type": ch.get("type"), "title": ch.get("title"),
                                "at": A.a1(ch["anchor"].get("r", 0), ch["anchor"].get("c", 0)),
                                "data": [s.get("val") for s in ch.get("series", [])][:6]} for ch in sh["charts"]]}
            if mr >= 0:
                top = min(r for r, _ in sh["cells"])
                info["headerRow"] = top + 1
                info["header"] = [display_value(doc, sh, top, c, sh["cells"].get((top, c)) or {}) for c in range(0, min(mc, 29) + 1)]
                types = []
                for c in range(0, min(mc, 29) + 1):
                    cnt = {}
                    for r in range(top + 1, min(mr, top + 1000) + 1):
                        cell = sh["cells"].get((r, c))
                        if not cell:
                            continue
                        v = cell_value(doc, sh, r, c, cell)
                        nf = (cell.get("s") or {}).get("nf", "")
                        t = ("formula:" if cell.get("f") else "") + (
                            "date" if isinstance(v, (int, float)) and re.search(r"[ymd]", nf) else
                            "number" if isinstance(v, (int, float)) and not isinstance(v, bool) else
                            "bool" if isinstance(v, bool) else "text" if isinstance(v, str) else "empty")
                        cnt[t] = cnt.get(t, 0) + 1
                    types.append(max(cnt, key=cnt.get) if cnt else "empty")
                info["columnTypes"] = {A.col_name(c): t for c, t in enumerate(types)}
            out["sheets"].append(info)
    return out


@route("GET", r"/api/docs/([\w-]+)/changes")
@_wrap
def r_changes(h, m, q, body):
    doc = _doc(m)
    since = int((q.get("since") or ["0"])[0])
    actor = (q.get("actor") or [None])[0]
    names = {sh["id"]: sh["name"] for sh in doc.sheets}
    out = []
    with doc.lock:
        for e in doc.log:
            if e["v"] <= since or (actor and e["actor"] != actor):
                continue
            ranges, metas = [], []
            for op in e["ops"]:
                if op["op"] == "setCells" and op["cells"]:
                    rs = [x[0] for x in op["cells"]]
                    cs = [x[1] for x in op["cells"]]
                    ranges.append(f"{names.get(op['sheet'], op['sheet'])}!{A.range_a1(min(rs), min(cs), max(rs), max(cs))}"
                                  f" ({len(op['cells'])}셀)")
                elif op["op"] == "setSheetMeta":
                    metas.append(f"{names.get(op['sheet'], op['sheet'])}: {', '.join(sorted(op['meta']))}")
                elif op["op"] == "sheets":
                    metas.append("시트 구성: " + ", ".join(names.get(i, i) for i in op["order"]))
            out.append({"v": e["v"], "actor": e["actor"], "label": e["label"], "batchId": e["batchId"],
                        "time": time.strftime("%H:%M:%S", time.localtime(e["ts"])), "cells": ranges, "meta": metas})
    return {"version": doc.version, "dirty": doc.dirty, "changes": out}


@route("GET", r"/api/docs/([\w-]+)/selection")
def r_selection(h, m, q, body):
    doc = _doc(m)
    names = {sh["id"]: sh["name"] for sh in doc.sheets}
    with doc.lock:
        items = sorted(doc.presence.values(), key=lambda p: -p["ts"])
    return {"selection": [{"sheet": names.get(p.get("sheet"), p.get("sheet")), "range": p.get("sel"),
                           "active": p.get("active"), "secondsAgo": round(time.time() - p["ts"])} for p in items
                          if p.get("actor") == "user"]}
