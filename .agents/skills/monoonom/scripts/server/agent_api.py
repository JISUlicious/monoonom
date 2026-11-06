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
from . import pivots as PV
from .app import SKILL_DIR, ApiError, route
from .docs_api import REGISTRY, _doc, _wrap
from . import refs
from . import region as RG
from .session import new_id

STYLE_KEYS = {"b", "i", "u", "st", "fs", "ff", "fc", "bg", "ha", "va", "wrap", "ind", "rot", "nf",
              "bl", "br", "bt", "bb"}
EXCEL_EPOCH = dt.datetime(1899, 12, 30)
ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


# ------------------------------------------------------------------ helpers

def resolve(doc, ref, sheet=None):
    """'Sheet!A1:B2' | 'A1:B2' (+ optional sheet) -> (sheet dict, (r1,c1,r2,c2))."""
    if not ref:
        raise ApiError(400, "A range is required.")
    sname, rng = A.split_sheet(ref)
    sh = doc.sheet(sname or sheet) if (sname or sheet) else doc.sheets[min(doc.active_sheet, len(doc.sheets) - 1)]
    try:
        return sh, A.parse_range(rng)
    except ValueError as e:
        raise ApiError(400, str(e))


def parse_at(doc, at):
    """'E2' or 'Sheet!E2' -> (sheet dict or None, r, c); bad input is a 400, not a server error."""
    sname, rng = A.split_sheet(at)
    try:
        r, c = A.parse_range(rng)[:2]
    except ValueError as e:
        raise ApiError(400, f"Invalid chart position (--at): {at} ({e})")
    return (doc.sheet(sname) if sname else None), r, c


def find_chart(doc, meta, meta_of, cid=None, sheet=None, index=None):
    """(sheet, charts list, index) of a chart. Chart ids are unique in the workbook, so without --sheet
    every sheet is searched (`chart ls` lists all sheets); an index needs a sheet (default: active)."""
    if sheet:
        sheets = [doc.sheet(sheet)]
    elif cid:
        sheets = list(doc.sheets)
    else:
        sheets = [doc.sheets[min(doc.active_sheet, len(doc.sheets) - 1)]]
    for sh in sheets:
        # look without registering a change for every sheet searched (that would emit no-op sheet ops)
        current = meta.get(sh["id"], {}).get("charts", sh["charts"])
        for i, ch in enumerate(current):
            if (cid and ch.get("id") == cid) or (not cid and index is not None and i == int(index)):
                return sh, meta_of(sh, "charts"), i
    where = f"sheet '{sheets[0]['name']}'" if len(sheets) == 1 else "all sheets"
    raise ApiError(404, f"Chart not found: {cid or f'#{index}'} (searched {where}). Chart ids change when the file is reopened or "
                        "the server restarts. Get the current ids with `xl.py chart ls`.")


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


def translate(formula: str, dr: int, dc: int) -> str:
    """Copy/fill semantics (same rewriter as the browser, see refs.py)."""
    return refs.translate(formula, dr, dc)


def rename_in_formula(formula, old, new):
    lc = old.lower()
    return refs.rewrite_refs(formula, lambda r: {**r, "sheet": new} if r["sheet"] is not None and r["sheet"].lower() == lc else r)


def merged_style(cur, patch=None, unset=()):
    st = dict(cur or {})
    for k, v in (patch or {}).items():
        if k not in STYLE_KEYS:
            raise ApiError(400, f"Unknown format key: {k} (available: {', '.join(sorted(STYLE_KEYS))})")
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
        return {"status": "stale", "source": doc.calc_source, "note": f"Node calculation failed: {err}"}
    return {"status": "stale", "source": doc.calc_source,
            "note": "Formula results may be out of date. Open the preview or install Node.js 18+ to compute them."}


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

    from types import SimpleNamespace
    from . import agent_more
    ctx = SimpleNamespace(doc=doc, cur=cur, put=put, meta=meta, meta_of=meta_of, cellbuf=cellbuf, ops=ops,
                          summary=summary, structural=False, transformable=True,
                          sheet_of=lambda e: doc.sheet(e["sheet"]) if e.get("sheet") else doc.sheets[min(doc.active_sheet, len(doc.sheets) - 1)])

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
                raise ApiError(400, f"The values ({h}x{w}) do not match the range ({A.range_a1(r1, c1, r2, c2)}).")
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
                raise ApiError(400, "At most 200,000 cells can be written at once.")
            for i in range(r2 - r1 + 1):
                for j in range(c2 - c1 + 1):
                    raw = vals[0][0] if fill else (vals[i][j] if j < len(vals[i]) else None)
                    fields, nf = parse_value(raw, parse)
                    if fill and "f" in fields:
                        fields["f"] = translate(fields["f"], i, j)
                    put(sh, r1 + i, c1 + j, cell_with(cur(sh, r1 + i, c1 + j), fields, nf))
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} write")
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
                    elif what == "comments":
                        put(sh, r, c, {k: v for k, v in cell.items() if k != "cm"} or None)
                    else:
                        keep = {k: v for k, v in cell.items() if k in ("s", "cm")}
                        put(sh, r, c, keep or None)
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} clear ({what})")
        elif kind == "format":
            sh, g = resolve(doc, e.get("range"), sheet)
            r1, c1, r2, c2 = clamp(sh, g, pending=cellbuf.get(sh["id"]))
            patch, unset = dict(e.get("set") or {}), e.get("unset") or []
            border = patch.pop("border", None) or e.get("border")
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
                raise ApiError(400, "Formats can be applied to at most 200,000 cells at once.")
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
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} format")
        elif kind in ("merge", "unmerge"):
            sh, (r1, c1, r2, c2) = resolve(doc, e.get("range"), sheet)
            merges = meta_of(sh, "merges")
            keep = []
            for mref in merges:
                a, b, c_, d = A.parse_range(mref)
                if not (a <= r2 and r1 <= c_ and b <= c2 and c1 <= d):
                    keep.append(mref)
            regions = [(rr, c1, rr, c2) for rr in range(r1, r2 + 1)] if e.get("across") else [(r1, c1, r2, c2)]
            if kind == "merge" and e.get("across"):
                for (a_, b_, c_, d_) in regions:
                    if b_ != d_:
                        keep.append(A.range_a1(a_, b_, c_, d_))
                        for cc in range(b_ + 1, d_ + 1):
                            if cur(sh, a_, cc):
                                put(sh, a_, cc, {"s": cur(sh, a_, cc)["s"]} if cur(sh, a_, cc).get("s") else None)
            elif kind == "merge":
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
            summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} {'merge' if kind == 'merge' else 'unmerge'}")
        elif kind in ("colWidth", "rowHeight", "hide", "unhide"):
            sh, (r1, c1, r2, c2) = resolve(doc, e.get("range") or e.get("cols") or e.get("rows"), sheet)
            axis = "cols" if (kind == "colWidth" or (kind in ("hide", "unhide") and r1 == 0 and r2 == A.MAX_ROWS - 1)) else "rows"
            dims = meta_of(sh, axis)
            lo, hi = (c1, c2) if axis == "cols" else (r1, r2)
            if hi - lo > 5000:
                raise ApiError(400, "At most 5,000 rows / columns can be changed at once.")
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
            summary.append(f"{sh['name']} freeze panes {cellref or 'off'}")
        elif kind == "filter":
            sh, g = (resolve(doc, e["range"], sheet) if e.get("range") else (doc.sheet(sheet) if sheet else doc.sheets[doc.active_sheet], None))
            if g:
                g, auto = RG.auto_range(sh, g, cur)
                r1, c1, r2, c2 = clamp(sh, g, pending=cellbuf.get(sh["id"]))
                meta.setdefault(sh["id"], {})["autoFilter"] = {"ref": A.range_a1(r1, c1, max(r2, r1 + 1), c2), "cols": {}}
            else:
                meta.setdefault(sh["id"], {})["autoFilter"] = None
            meta[sh["id"]]["filterHidden"] = []
            summary.append(f"{sh['name']} autofilter {'on' if g else 'off'}")
        elif kind == "chart":
            sh_data, g = resolve(doc, e.get("data"), sheet)
            g, _ = RG.auto_range(sh_data, g, cur)
            at_sheet = parse_at(doc, e["at"])[0] if e.get("at") else None
            if at_sheet and e.get("sheet") and doc.sheet(e["sheet"]) is not at_sheet:
                raise ApiError(400, "--sheet and --at name different sheets.")
            target = doc.sheet(e["sheet"]) if e.get("sheet") else at_sheet or sh_data
            chart = chart_from_range(doc, sh_data, g, e, cellbuf.get(sh_data["id"], {}))
            charts = meta_of(target, "charts")
            charts.append(chart)
            summary.append(f"{target['name']} add chart ({chart['type']}, id={chart['id']})")
        elif kind == "deleteChart":
            sh, charts, idx = find_chart(doc, meta, meta_of, e.get("id"), sheet, e.get("index"))
            meta[sh["id"]]["charts"] = [ch for i, ch in enumerate(charts) if i != idx]
            summary.append(f"{sh['name']} delete chart")
        elif kind == "addSheet":
            name = (e.get("name") or "").strip() or _unique_name(doc, "Sheet")
            _check_sheet_name(doc, name)
            sid = new_id("s")
            order = [s["id"] for s in doc.sheets]
            at = int(e.get("at", len(order)))
            order.insert(max(0, min(at, len(order))), sid)
            ops.append({"op": "sheets", "order": order, "add": {sid: {"name": name, "cells": []}}})
            structural = True
            summary.append(f"add sheet '{name}'")
        elif kind == "renameSheet":
            sh = doc.sheet(sheet or e.get("from"))
            name = (e.get("name") or "").strip()
            _check_sheet_name(doc, name, exclude=sh)
            old = sh["name"]
            meta.setdefault(sh["id"], {})["name"] = name
            for other in doc.sheets:
                for (r, c), cell in other["cells"].items():
                    if cell.get("f") and old.lower() in cell["f"].lower():
                        nf = rename_in_formula(cell["f"], old, name)
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
                                    s[k] = rename_in_formula("=" + s[k], old, name)[1:]
                        ch["dirty"] = True
                    meta[other["id"]]["charts"] = new
                pivots = copy.deepcopy(meta_of(other, "pivots") or [])
                if any(p.get("source") and old.lower() in p["source"].lower() for p in pivots):
                    for p in pivots:
                        if p.get("source"):
                            p["source"] = rename_in_formula("=" + p["source"], old, name)[1:]
                    meta[other["id"]]["pivots"] = pivots
            names = {k: rename_in_formula(v if v.startswith("=") else "=" + v, old, name).lstrip("=")
                     for k, v in doc.defined_names.items()}
            ops.append({"op": "setWorkbook", "props": {"definedNames": names}})
            structural = True
            summary.append(f"rename sheet '{old}' → '{name}'")
        elif kind == "deleteSheet":
            sh = doc.sheet(sheet or e.get("name"))
            if len(doc.sheets) <= 1:
                raise ApiError(400, "The last sheet cannot be deleted.")
            ops.append({"op": "sheets", "order": [s["id"] for s in doc.sheets if s is not sh], "add": {}})
            structural = True
            ctx.transformable = False
            summary.append(f"delete sheet '{sh['name']}'")
        elif kind == "name":
            names = dict(doc.defined_names)
            if e.get("ref"):
                names[e["name"]] = e["ref"].lstrip("=")
            else:
                names.pop(e["name"], None)
            ops.append({"op": "setWorkbook", "props": {"definedNames": names}})
            summary.append(f"define name {e['name']}")
        elif kind in agent_more.HANDLERS:
            agent_more.HANDLERS[kind](ctx, e)
        else:
            raise ApiError(400, f"Unknown edit op: {kind} (see references/edits.md)")

    for sid, cells in cellbuf.items():
        ops.append({"op": "setCells", "sheet": sid, "cells": [[r, c, cell] for (r, c), cell in sorted(cells.items())]})
    for sid, m in meta.items():
        ops.append({"op": "setSheetMeta", "sheet": sid, "meta": m})
    # structure ops last so cell/meta ops still find their sheets (deleted sheets are dropped)
    ops.sort(key=lambda o: o["op"] == "sheets")
    build_ops.last_transformable = ctx.transformable
    return ops, structural or ctx.structural, summary


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
        raise ApiError(400, f"Border kinds: {', '.join(sel)} or none")
    for k, on in zip(("bt", "bb", "bl", "br"), sel[kind]):
        if on:
            st[k] = side
    return st


def _check_sheet_name(doc, name, exclude=None):
    if not name or len(name) > 31 or re.search(r"[\\/?*\[\]:]", name):
        raise ApiError(400, "Sheet names are 1–31 characters without \\ / ? * [ ] :")
    if any(s is not exclude and s["name"].lower() == name.lower() for s in doc.sheets):
        raise ApiError(409, f"A sheet with this name exists: {name}")


def _unique_name(doc, base):
    i = 1
    names = {s["name"].lower() for s in doc.sheets}
    while f"{base}{i}".lower() in names:
        i += 1
    return f"{base}{i}"


def _lit_label(xs):
    """A series' own list of values (no cell range), as `chart ls` shows it."""
    if not xs:
        return None
    shown = ", ".join("" if x is None else (f"{x:g}" if isinstance(x, float) else str(x)) for x in xs[:3])
    return f"{{{shown}{', …' if len(xs) > 3 else ''}}} ({len(xs)} values stored in the chart)"


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
    if ctype == "scatter" and len(series) > 1 and not header_col:
        x = series.pop(0)  # numeric first column = X values of every series (Excel; same as the UI's fromRange)
        for s in series:
            s["cat"] = x["val"]
    preset = {"barH": ("bar", "bar", "clustered"), "barStacked": ("bar", "col", "stacked")}.get(ctype)
    t, bar_dir, grouping = preset or (ctype, "col", "standard" if ctype in ("line", "area") else "clustered")
    if t not in ("bar", "line", "area", "pie", "doughnut", "scatter", "radar"):
        raise ApiError(400, "Chart types: bar, barH, barStacked, line, area, pie, doughnut, scatter, radar")
    at = e.get("at")
    ar, ac = (parse_at(doc, at)[1:] if at else (r1, min(c2 + 2, A.MAX_COLS - 10)))
    return {"id": new_id("ch"), "type": t, "barDir": bar_dir, "grouping": grouping, "title": e.get("title"),
            "legend": e.get("legend", "r" if (len(series) > 1 or t in ("pie", "doughnut")) else None),
            "xTitle": e.get("xTitle"), "yTitle": e.get("yTitle"), "series": series, "dirty": True,
            "anchor": {"c": ac, "r": ar, "dx": 0, "dy": 0, "w": int(e.get("width", 480)), "h": int(e.get("height", 288))}}


# ------------------------------------------------------------------ structural edits (P3)

STRUCTURAL = {"insertRows", "deleteRows", "insertCols", "deleteCols", "sort"}


def _axis_pos(spec, axis):
    """'5' / 5 (1-based row) or 'C' (column) -> 0-based index; '5:7' / 'C:D' -> (start, count)."""
    t = str(spec).strip().replace("$", "")
    a, _, b = t.partition(":")
    conv = (lambda x: int(x) - 1) if axis == "row" else (lambda x: A.col_index(x) if x.isalpha() else int(x) - 1)
    try:
        lo = conv(a)
        hi = conv(b) if b else lo
    except ValueError:
        raise ApiError(400, f"Invalid {'row' if axis == 'row' else 'column'} position: {spec}")
    if lo < 0 or hi < lo:
        raise ApiError(400, f"Invalid position: {spec}")
    return lo, hi - lo + 1


def shift_ops(doc, sh, axis, at, count):
    """State ops for inserting (count>0) / deleting (count<0) rows/cols, mirroring Workbook.shift()."""
    problem = PV.shift_problem(sh, axis, at, count)
    if problem:
        raise ApiError(409, problem)
    is_row = axis == "row"
    k = 0 if is_row else 1
    target = sh["name"]
    new_cells = {}
    for (r, c), cell in sh["cells"].items():
        p = (r, c)[k]
        if p < at:
            new_cells[(r, c)] = cell
            continue
        if count < 0 and p < at - count:
            continue  # deleted
        np_ = p + count
        if np_ >= (A.MAX_ROWS if is_row else A.MAX_COLS):
            continue
        new_cells[(np_, c) if is_row else (r, np_)] = cell
    ops = []
    # formulas everywhere (including moved cells) follow the shift
    for other in doc.sheets:
        cells = new_cells if other is sh else other["cells"]
        changed = {}
        for key, cell in cells.items():
            if cell.get("f"):
                nf = refs.adjust(cell["f"], other["name"], target, axis, at, count)
                if nf != cell["f"]:
                    changed[key] = dict(cell, f=nf)
        if other is sh:
            new_cells.update(changed)
            old_keys = set(sh["cells"])
            out = []
            for key in old_keys | set(new_cells):
                before, after = sh["cells"].get(key), new_cells.get(key)
                if before != after:
                    out.append([key[0], key[1], after])
            if out:
                ops.append({"op": "setCells", "sheet": sh["id"], "cells": sorted(out, key=lambda x: (x[0], x[1]))})
        elif changed:
            ops.append({"op": "setCells", "sheet": other["id"], "cells": [[r, c, v] for (r, c), v in sorted(changed.items())]})

    def shift_index(i):
        if i < at:
            return i
        if count > 0:
            return i + count
        return i + count if i >= at - count else None

    def shift_range(ref):
        r1, c1, r2, c2 = A.parse_range(ref)
        lo, hi = (r1, r2) if is_row else (c1, c2)
        if count > 0:
            lo = lo + count if lo >= at else lo
            hi = hi + count if hi >= at else hi
        else:
            end = at - count - 1
            if lo >= at and hi <= end:
                return None
            lo = lo + count if lo > end else (at if lo >= at else lo)
            hi = hi + count if hi > end else (at - 1 if hi >= at else hi)
            if hi < lo:
                return None
        return A.range_a1(lo, c1, hi, c2) if is_row else A.range_a1(r1, lo, r2, hi)

    meta = {}
    dims_key = "rows" if is_row else "cols"
    dims = {}
    for key, e in sh[dims_key].items():
        ni = shift_index(int(key))
        if ni is not None:
            dims[str(ni)] = e
    meta[dims_key] = dims
    if is_row:
        meta["filterHidden"] = sorted(x for x in (shift_index(r) for r in sh["filterHidden"]) if x is not None)
    merges = []
    for mref in sh["merges"]:
        nr = shift_range(mref)
        if nr and ":" in nr:
            merges.append(nr)
    meta["merges"] = merges
    af = sh.get("autoFilter")
    if af and af.get("ref"):
        nr = shift_range(af["ref"])
        meta["autoFilter"] = {**af, "ref": nr} if nr else None
        if not nr:
            meta["filterHidden"] = []
    charts = copy.deepcopy(sh["charts"])
    for ch in charts:
        p = "r" if is_row else "c"
        if ch["anchor"].get(p, 0) >= at:
            ch["anchor"][p] = max(at, ch["anchor"][p] + count)
    meta["charts"] = charts
    pivots = copy.deepcopy(sh.get("pivots") or [])
    for p in pivots:  # the pivot moves as a whole (cutting through it was refused above)
        nr = shift_range(p["ref"])
        if nr:
            p["ref"] = nr
    meta["pivots"] = pivots
    for other in doc.sheets:  # pivot sources on the shifted sheet (any sheet's pivots)
        pvs = pivots if other is sh else copy.deepcopy(other.get("pivots") or [])
        touched = False
        for p in pvs:
            if p.get("source"):
                n = refs.adjust("=" + p["source"], other["name"], target, axis, at, count)[1:]
                if n != p["source"]:
                    p["source"], touched = n, True
        if other is not sh and touched:
            ops.append({"op": "setSheetMeta", "sheet": other["id"], "meta": {"pivots": pvs}})
    for other in doc.sheets:  # chart series referencing the shifted sheet (any sheet)
        chs = charts if other is sh else copy.deepcopy(other["charts"])
        touched = False
        for ch in chs:
            for se in ch.get("series", []):
                for key in ("val", "cat", "nameRef"):
                    if se.get(key):
                        n = refs.adjust("=" + se[key], other["name"], target, axis, at, count)[1:]
                        if n != se[key]:
                            se[key], ch["dirty"], touched = n, True, True
        if other is not sh and touched:
            ops.append({"op": "setSheetMeta", "sheet": other["id"], "meta": {"charts": chs}})
    ops.append({"op": "setSheetMeta", "sheet": sh["id"], "meta": meta})
    names = {n: refs.adjust("=" + v.lstrip("="), target, target, axis, at, count)[1:] if "!" in v else v
             for n, v in doc.defined_names.items()}
    if names != doc.defined_names:
        ops.append({"op": "setWorkbook", "props": {"definedNames": names}})
    return ops


_RANK = lambda v: 4 if v is None or v == "" else 0 if isinstance(v, (int, float)) and not isinstance(v, bool) else 1 if isinstance(v, str) else 2 if isinstance(v, bool) else 3


def sort_ops(doc, sh, g, keys, header):
    """Sort rows of a range like the UI (visible rows only, blanks last, formulas follow their row)."""
    r1, c1, r2, c2 = clamp(sh, g)
    pv = PV.overlapping(sh, (r1, c1, r2, c2))
    if pv:
        raise ApiError(409, f"Cannot sort a range that contains pivot table '{pv[0]['name']}' ({pv[0]['ref']}).")
    if header is None:
        top = [cell_value(doc, sh, r1, c, sh["cells"].get((r1, c)) or {}) for c in range(c1, c2 + 1)]
        nxt = [cell_value(doc, sh, r1 + 1, c, sh["cells"].get((r1 + 1, c)) or {}) for c in range(c1, c2 + 1)]
        header = any(isinstance(v, str) for v in top) and (any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in nxt) or all(isinstance(v, str) for v in top))
    start = r1 + (1 if header else 0)
    hidden = set(sh["filterHidden"])
    rows = [r for r in range(start, r2 + 1) if r not in hidden]

    def sort_key(r):
        out = []
        for kc, asc in keys:
            v = cell_value(doc, sh, r, kc, sh["cells"].get((r, kc)) or {})
            rank = _RANK(v)
            if rank == 4:
                out.append((1, 0, 0))  # blanks last in both directions
                continue
            val = v.casefold() if isinstance(v, str) else (float(v) if rank in (0, 2) else str(v))
            out.append((0, rank if asc else -rank, val if asc else _Neg(val)))
        return out + [r]

    order = sorted(rows, key=sort_key)
    cells = []
    for dst, src in zip(rows, order):
        for c in range(c1, c2 + 1):
            cell = sh["cells"].get((src, c))
            if cell and cell.get("f") and dst != src:
                cell = dict(cell, f=refs.translate(cell["f"], dst - src, 0))
            if sh["cells"].get((dst, c)) != cell:
                cells.append([dst, c, cell])
    return ([{"op": "setCells", "sheet": sh["id"], "cells": cells}] if cells else []), header, A.range_a1(start, c1, r2, c2)


class _Neg:
    """Reverses ordering for descending sort keys of any comparable type."""
    __slots__ = ("v",)

    def __init__(self, v):
        self.v = v

    def __lt__(self, o):
        return o.v < self.v

    def __eq__(self, o):
        return self.v == o.v


def structural_batch(doc, e):
    """One structural edit -> (ops, shifts, transformable, summary)."""
    kind = e["op"]
    sheet = e.get("sheet")
    if kind == "sort":
        sh, g = resolve(doc, e.get("range"), sheet)
        af = sh.get("autoFilter")
        if af and af.get("ref") and (g[0], g[1]) == (g[2], g[3]):
            fr = A.parse_range(af["ref"])
            if fr[0] <= g[0] <= fr[2] and fr[1] <= g[1] <= fr[3]:
                g = RG.table_body(sh, fr)  # inside the filter: sort the filter's table (header kept)
                e = {**e, "header": True}
        if (g[0], g[1]) == (g[2], g[3]) or g[2] - g[0] >= A.MAX_ROWS - 1 or g[3] - g[1] >= A.MAX_COLS - 1:
            g, _ = RG.auto_range(sh, g)
        keys = []
        for k in e.get("by") or []:
            if isinstance(k, str):
                col, _, d = k.partition(":")
                k = {"col": col, "asc": d.lower() not in ("desc", "d")}
            col = k["col"]
            keys.append((A.col_index(col) if str(col).isalpha() else int(col) - 1, bool(k.get("asc", True))))
        if not keys:
            raise ApiError(400, "Sort keys (by) are required, e.g. [\"C:desc\"]")
        ops, header, rng = sort_ops(doc, sh, g, keys, e.get("header"))
        return ops, [], False, f"{sh['name']}!{rng} sort ({', '.join(A.col_name(c) + ('↑' if a else '↓') for c, a in keys)})" + (" excluding header" if header else "")
    axis = "row" if kind.endswith("Rows") else "col"
    sh = doc.sheet(sheet) if sheet else doc.sheets[doc.active_sheet]
    if "range" in e and "!" in str(e["range"]):
        sname, rest = A.split_sheet(e["range"])
        sh, e = doc.sheet(sname), {**e, "at": rest}
    at, n = _axis_pos(e.get("at", e.get("range")), axis)
    count = int(e.get("count", n))
    if count <= 0:
        raise ApiError(400, "count must be 1 or more.")
    if count > (5000 if axis == "col" else 100000):
        raise ApiError(400, "Too many rows / columns to insert or delete at once.")
    signed = count if kind.startswith("insert") else -count
    ops = shift_ops(doc, sh, axis, at, signed)
    what = f"{at + 1}" if axis == "row" else A.col_name(at)
    label = f"{sh['name']} {'insert' if signed > 0 else 'delete'} {'rows' if axis == 'row' else 'columns'} {what}{' (' + str(count) + ')' if count > 1 else ''}"
    return ops, [{"sheet": sh["id"], "axis": axis, "at": at, "count": signed}], True, label


@route("POST", r"/api/docs/([\w-]+)/edit")
@_wrap
def r_edit(h, m, q, body):
    """Apply agent edits. Consecutive non-structural edits form one batch; each row/column insert/delete
    or sort is its own batch (it changes coordinates the following edits refer to)."""
    doc = _doc(m)
    edits = body.get("edits") or body.get("ops") or []
    if isinstance(edits, dict):
        edits = [edits]
    groups, cur = [], []
    for e in edits:
        if e.get("op") in STRUCTURAL:
            if cur:
                groups.append(cur)
                cur = []
            groups.append([e])
        else:
            cur.append(e)
    if cur:
        groups.append(cur)
    results = []
    with doc.lock:
        for gi, group in enumerate(groups):
            batch = (body.get("batchId") if len(groups) == 1 else None) or new_id("b")
            if group[0].get("op") in STRUCTURAL:
                ops, shifts, transformable, label = structural_batch(doc, group[0])
                summary, structural = [label], True
            else:
                ops, structural, summary = build_ops(doc, group)
                shifts, transformable = [], build_ops.last_transformable
            if not ops:
                results.append({"version": doc.version, "batchId": None, "summary": summary or []})
                continue
            v = doc.apply(ops, actor="agent", client_id=body.get("clientId", "agent-cli"), batch_id=batch,
                          label=(body.get("label") if len(groups) == 1 else None) or "; ".join(summary)[:120],
                          structural=structural, shifts=shifts, transformable=transformable)
            results.append({"version": v, "batchId": batch, "summary": summary})
    last = results[-1] if results else {"version": doc.version, "batchId": None, "summary": []}
    return {**last, "batches": results, "summary": [x for r in results for x in r["summary"]]}


@route("POST", r"/api/docs/([\w-]+)/pivots/refresh")
@_wrap
def r_pivot_refresh(h, m, q, body):
    """Recompute pivot tables' data cells from their sources (all pivots, or one by name / sheet),
    in the layout the file has. One undoable batch; the file's pivot cache is still rebuilt by Excel
    when it opens the saved file (refreshOnLoad)."""
    doc = _doc(m)
    calc = ensure_calc(doc)
    name, sheet = (body.get("name") or "").strip().lower(), (body.get("sheet") or "").strip().lower()
    with doc.lock:
        targets = [(sh, p) for sh in doc.sheets for p in sh.get("pivots") or []
                   if (not name or p["name"].lower() == name) and (not sheet or sh["name"].lower() == sheet)]
        if not targets:
            raise ApiError(404, f"Pivot table not found: {body.get('name') or body.get('sheet') or '(none in this document)'}")
        value_at = lambda s, r, c: cell_value(doc, s, r, c, s["cells"].get((r, c)) or {})  # noqa: E731
        ops, done, notes = [], [], []
        for sh, p in targets:
            cells, n = PV.refresh(doc, sh, p, value_at)
            notes += n
            if cells:
                ops.append({"op": "setCells", "sheet": sh["id"], "cells": cells})
            done.append({"sheet": sh["name"], "name": p["name"], "range": p["ref"], "changed": len(cells)})
        if not ops:
            return {"version": doc.version, "batchId": None, "pivots": done, "notes": notes, "calc": calc}
        batch = new_id("b")
        v = doc.apply(ops, actor=body.get("actor", "agent"), client_id=body.get("clientId", "agent-cli"), batch_id=batch,
                      label="refresh pivot: " + ", ".join(p["name"] for _, p in targets), pivot_write=True)
    return {"version": v, "batchId": batch, "pivots": done, "notes": notes, "calc": calc}


@route("POST", r"/api/docs/([\w-]+)/pivots")
@_wrap
def r_pivot_create(h, m, q, body):
    """Create a pivot table: body {source, rows: [field], cols: [field], values: [{field, fn, name?}],
    filters: [field], at: 'Sheet!A3' (a missing sheet is created), name?}. Fields are the source's
    header names. Written in tabular form (pivots.build); saving writes a real pivot definition and
    cache (xlsx_io._create_pivot) that Excel rebuilds when it opens the file."""
    doc = _doc(m)
    calc = ensure_calc(doc)
    with doc.lock:
        cur = doc.sheets[min(doc.active_sheet, len(doc.sheets) - 1)]
        ssh, g = resolve(doc, body.get("source"), cur["name"])
        g, _ = RG.auto_range(ssh, g, None, trim_totals=True)  # one cell / whole columns -> the table there
        r1, c1, r2, c2 = clamp(ssh, g)
        if r2 <= r1:
            raise ApiError(400, "The source range needs a header row and data rows.")
        val = lambda r, c: cell_value(doc, ssh, r, c, ssh["cells"].get((r, c)) or {})  # noqa: E731
        header = [val(r1, c) for c in range(c1, c2 + 1)]
        if any(hv in (None, "") for hv in header):
            raise ApiError(400, "The source's first row (header) has an empty cell. Every column needs a name.")
        header = [str(hv) for hv in header]
        recs = [{header[j]: val(r, c1 + j) for j in range(len(header))} for r in range(r1 + 1, r2 + 1)]
        recs = [x for x in recs if any(v not in (None, "") for v in x.values())]
        parse_vals = []
        for v in body.get("values") or []:
            parse_vals.append(dict(v) if isinstance(v, dict) else dict(zip(("field", "fn"), (str(v).split(":", 1) + ["sum"])[:2])))
        try:
            L = PV.labels(body.get("lang"))
            built = PV.build(recs, header, list(body.get("rows") or []), list(body.get("cols") or []), parse_vals, list(body.get("filters") or []), body.get("lang"))
        except ValueError as e:
            raise ApiError(400, str(e))
        at = body.get("at") or ""
        sname, cell = A.split_sheet(at) if at else (None, "A3")
        try:
            tr, tc = A.parse_range(cell or "A3")[:2]
        except ValueError as e:
            raise ApiError(400, f"Invalid pivot position (--at): {at} ({e})")
        if not at:
            sname = None
        ops, structural = [], False
        tsh = None
        if sname:
            tsh = next((s for s in doc.sheets if s["name"].lower() == sname.lower()), None)
        new_sheet = tsh is None
        if new_sheet:
            sname = sname or _unique_name(doc, L["sheet"])
            _check_sheet_name(doc, sname)
            sid = new_id("s")
            sj = {"id": sid, "name": sname, "cells": [], "charts": [], "pivots": [], "merges": [], "cols": {}, "rows": {}}
            order = [x["id"] for x in doc.sheets]
            order.insert(order.index(ssh["id"]) + 1 if ssh["id"] in order else len(order), sid)
            ops.append({"op": "sheets", "order": order, "add": {sid: sj}})
            structural = True
        grid = built["grid"]
        height, width = len(grid), max(len(x) for x in grid)
        area = (tr, tc, tr + height - 1, tc + width - 1)
        if tsh is not None:
            busy = [k for k in tsh["cells"] if area[0] <= k[0] <= area[2] and area[1] <= k[1] <= area[3]
                    and (tsh["cells"][k].get("v") not in (None, "") or tsh["cells"][k].get("f"))]
            if busy or PV.overlapping(tsh, area):
                raise ApiError(409, f"The pivot's area {A.range_a1(*area)} has data. Choose an empty place or a new sheet (--at 'NewSheet!A3').")
        names = {p["name"].lower() for x in doc.sheets for p in x.get("pivots") or []}
        name = (body.get("name") or "").strip() or next(L["table"].format(n=i) for i in range(1, 999) if L["table"].format(n=i).lower() not in names)
        if name.lower() in names:
            raise ApiError(409, f"A pivot table with this name exists: {name}")
        top = built["table_top"]
        bold = {"b": True}
        header_rows = built["layout"]["firstDataRow"]
        cells = []
        for i, line in enumerate(grid):
            for j, v in enumerate(line):
                if v is None:
                    continue
                cell = {"v": v}
                n_rows = len(body.get("rows") or [])
                is_header = top <= i < top + header_rows
                is_total = any(PV.is_total_label(x) for x in line[:n_rows])
                if is_header or is_total or (i < top and j == 0):
                    cell["s"] = dict(bold)
                cells.append([tr + i, tc + j, cell])
        source = f"{A.quote_sheet(ssh['name'])}!{A.range_a1(r1, c1, r2, c2)}"
        pm = {"name": name, "ref": A.range_a1(tr + top, tc, tr + height - 1, tc + width - 1), "source": source,
              "rows": list(body.get("rows") or []), "cols": list(body.get("cols") or []), "filters": list(body.get("filters") or []),
              "values": [{"name": v["name"], "field": v["field"], "fn": v["fn"]} for v in built["values"]],
              "layout": built["layout"], "create": True, "lang": "ko" if body.get("lang") == "ko" else "en"}
        sheets_now = [{"name": x["name"], "cells": [[r, c, it] for (r, c), it in x["cells"].items()]} for x in doc.sheets]
        from .xlsx_io import pivot_source_sig
        pm["srcSig"] = pivot_source_sig(sheets_now, source)
        target_id = tsh["id"] if tsh is not None else sid
        ops.append({"op": "setCells", "sheet": target_id, "cells": cells})
        ops.append({"op": "setSheetMeta", "sheet": target_id, "meta": {"pivots": (list(tsh.get("pivots") or []) if tsh else []) + [pm]}})
        batch = new_id("b")
        v = doc.apply(ops, actor=body.get("actor", "agent"), client_id=body.get("clientId", "agent-cli"), batch_id=batch,
                      label=f"create pivot table: {name}", structural=structural, pivot_write=True)
    return {"version": v, "batchId": batch, "name": name, "sheet": sname or tsh["name"], "range": pm["ref"],
            "area": A.range_a1(*area), "source": source, "calc": calc}


# ------------------------------------------------------------------ reads

@route("GET", r"/api/docs/([\w-]+)/range")
@_wrap
def r_range(h, m, q, body):
    doc = _doc(m)
    ref = (q.get("ref") or [""])[0]
    what = (q.get("what") or ["values"])[0]
    max_rows = int((q.get("max") or ["200"])[0])
    calc = ensure_calc(doc) if what != "formulas" else {"status": "skipped"}
    texts = None
    if what == "text":
        from .agent_more import display_texts
        sh0, g0 = resolve(doc, ref or "A1:Z50", (q.get("sheet") or [None])[0])
        a, b, c_, d_ = clamp(sh0, g0)
        texts = display_texts(doc, [(sh0["id"], a, b, min(c_, a + max_rows), min(d_, b + 99))])
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
                if what == "text":
                    row.append(texts.get((sh["id"], r, c), "") if cell else "")
                elif what == "values":
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


def unsupported_cells(sh):
    """Formula cells using functions the engine cannot calculate, and whether the file's original
    formula is still kept (it is written back unchanged on save while the formula is not edited)."""
    from .formula_compat import strip_prefixes, unsupported_functions
    out = []
    for (r, c), cell in sorted(sh["cells"].items()):
        f = cell.get("f")
        if not f:
            continue
        f = f if f.startswith("=") else "=" + f
        bad = unsupported_functions(f)
        if bad:
            fo = cell.get("fo")
            out.append({"cell": A.a1(r, c), "functions": list(bad), "kept": bool(fo and strip_prefixes(fo) == f)})
    return out


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
                    "unsupported": unsupported_cells(sh),
                    "autoFilter": (sh["autoFilter"] or {}).get("ref"),
                    "charts": [{"id": ch.get("id"), "type": ch.get("type"), "title": ch.get("title"),
                                **({"pivot": ch["pivot"]["table"]} if ch.get("pivot") else {}),
                                "at": A.a1(ch["anchor"].get("r", 0), ch["anchor"].get("c", 0)),
                                "data": [s.get("val") for s in ch.get("series", [])][:6],
                                # per series: name cell, X (categories / scatter x) and Y (values) ranges
                                "series": [{"name": s.get("name"), "nameRef": s.get("nameRef"), "x": s.get("cat") or _lit_label(s.get("catLit")), "y": s.get("val") or _lit_label(s.get("valLit")),
                                            **({"type": s["type"]} if s.get("type") else {})} for s in ch.get("series", [])]}
                               for ch in sh["charts"]],
                    "pivots": [{"name": p["name"], "range": p["ref"], "source": p.get("source"), "rows": p.get("rows", []),
                                "cols": p.get("cols", []), "values": [f"{v['fn']}({v['field']})" for v in p.get("values", [])],
                                "filters": p.get("filters", [])} for p in sh.get("pivots") or []]}
            if mr >= 0:
                r0, c0 = min(r for r, _ in sh["cells"]), min(c for _, c in sh["cells"])
                info["used"] = A.range_a1(r0, c0, mr, mc)
                tabs = [g for g in RG.tables(sh) if not PV.overlapping(sh, g)]  # a pivot's output is not a table
                info["tables"] = []
                for g in tabs:
                    body = RG.table_body(sh, g)
                    t = {"range": A.range_a1(*g), "headerRow": g[0] + 1,
                         "header": [display_value(doc, sh, g[0], c, sh["cells"].get((g[0], c)) or {}) for c in range(g[1], min(g[3], g[1] + 29) + 1)]}
                    if body[2] < g[2]:
                        t["totalRows"] = f"{body[2] + 2}" + (f":{g[2] + 1}" if g[2] > body[2] + 1 else "")
                    info["tables"].append(t)
                main = tabs[0] if tabs else (r0, c0, mr, mc)
                top = main[0]
                info["headerRow"] = top + 1
                info["header"] = [display_value(doc, sh, top, c, sh["cells"].get((top, c)) or {}) for c in range(main[1], min(main[3], main[1] + 29) + 1)]
                types = []
                body = RG.table_body(sh, main)
                for c in range(main[1], min(main[3], main[1] + 29) + 1):
                    cnt = {}
                    for r in range(top + 1, min(body[2], top + 1000) + 1):
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
                info["columnTypes"] = {A.col_name(main[1] + i): t for i, t in enumerate(types)}
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
                                  f" ({len(op['cells'])} cells)")
                elif op["op"] == "setSheetMeta":
                    metas.append(f"{names.get(op['sheet'], op['sheet'])}: {', '.join(sorted(op['meta']))}")
                elif op["op"] == "sheets":
                    metas.append("sheets: " + ", ".join(names.get(i, i) for i in op["order"]))
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


@route("GET", r"/api/docs/([\w-]+)/find")
@_wrap
def r_find(h, m, q, body):
    """Find cells by text (wildcards * ?), in values or formulas, on one sheet or all."""
    from .agent_more import _matcher
    doc = _doc(m)
    text = (q.get("q") or [""])[0]
    if not text:
        raise ApiError(400, "Text to find (q) is required.")
    where = (q.get("in") or ["values"])[0]
    rx = _matcher(text, (q.get("case") or ["0"])[0] == "1", (q.get("whole") or ["0"])[0] == "1")
    sheet = (q.get("sheet") or [None])[0]
    limit = int((q.get("max") or ["200"])[0])
    ensure_calc(doc)
    out = []
    with doc.lock:
        for sh in ([doc.sheet(sheet)] if sheet else doc.sheets):
            for (r, c), cell in sorted(sh["cells"].items()):
                t = cell.get("f") if where == "formulas" and cell.get("f") else display_value(doc, sh, r, c, cell)
                if t is not None and rx.search(str(t)):
                    out.append({"sheet": sh["name"], "cell": A.a1(r, c), "value": t})
                    if len(out) >= limit:
                        return {"matches": out, "truncated": True}
    return {"matches": out, "truncated": False}


@route("POST", r"/api/docs/([\w-]+)/goto")
@_wrap
def r_goto(h, m, q, body):
    """Move the user's preview to a sheet/range (not a document change)."""
    doc = _doc(m)
    sh, g = resolve(doc, body.get("range") or "A1", body.get("sheet"))
    doc.focused_at = time.time()
    doc.hub.publish("goto", {"docId": doc.id, "sheet": sh["id"], "range": A.range_a1(*g)})
    return {"ok": True, "sheet": sh["name"], "range": A.range_a1(*g), "clients": len(doc.hub)}


@route("POST", r"/api/docs/([\w-]+)/view")
@_wrap
def r_view(h, m, q, body):
    """Preview display settings: zoom (0.25-4) and formula display."""
    doc = _doc(m)
    data = {k: body[k] for k in ("zoom", "showFormulas", "follow") if k in body}
    doc.hub.publish("view", {"docId": doc.id, **data})
    return {"ok": True, **data, "clients": len(doc.hub)}
