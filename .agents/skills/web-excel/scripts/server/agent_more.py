"""Additional agent operations (P7) so everything the UI can do is also available to the agent.

Each handler takes a batch context `ctx` (current cell lookup, buffered writes, sheet-property
patches) from agent_api.build_ops and the edit dict. Logic mirrors the browser (filter.js,
app.js fill/copy/shiftCells/removeDuplicates/changeDecimals) so results are the same either way.
"""
from __future__ import annotations

import copy
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import unicodedata

from . import a1 as A
from . import refs
from .app import SKILL_DIR, ApiError
from .session import new_id

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


# ------------------------------------------------------------------ display text (what the grid shows)

def _is_date_nf(nf):
    clean = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", nf or "")
    return clean not in ("", "General") and (bool(re.search(r"[yd]", clean, re.I)) or bool(re.search(r"m{3,}", clean)) or bool(re.search(r"[hs]", clean, re.I)))


def format_py(v, nf):
    """Approximate number formatting when Node.js is not available (General, decimals, thousands,
    %, currency literals, dates). Node uses the exact browser formatter."""
    if v is None or v == "":
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dict):
        return v.get("err", "")
    if isinstance(v, str):
        return v
    nf = nf or "General"
    if _is_date_nf(nf) and 0 <= v < 2958466:
        d = (EXCEL_EPOCH + dt.timedelta(days=float(v))).replace(microsecond=0)
        clean = re.sub(r'"[^"]*"', "", nf)
        has_t = bool(re.search(r"[hs]", clean, re.I))
        has_d = bool(re.search(r"[yd]", clean, re.I))
        return d.strftime("%Y-%m-%d %H:%M:%S" if has_t and has_d else "%H:%M:%S" if has_t else "%Y-%m-%d")
    sec = nf.split(";")[0]
    lit_pre = "".join(re.findall(r'^"([^"]*)"', sec))
    lit_suf = "".join(re.findall(r'"([^"]*)"$', sec)) if not lit_pre else ""
    core = re.sub(r'"[^"]*"|_.|\*.|\[[^\]]*\]', "", sec)
    if nf == "General" or not re.search(r"[0#?]", core):
        s = f"{v:.10g}"
        return s.replace("e+", "E+").replace("e-", "E-")
    pct = core.count("%")
    x = v * (100 ** pct)
    dec = len(core.split(".")[1].rstrip("%").replace(",", "")) if "." in core else 0
    body = f"{abs(x):,.{dec}f}" if "," in core.split(".")[0] else f"{abs(x):.{dec}f}"
    return ("-" if x < 0 else "") + lit_pre + body + "%" * pct + lit_suf


def display_texts(doc, rects):
    """{(sheetId, r, c): text} for the given [(sheetId, r1, c1, r2, c2)] (Node exact, else approximate)."""
    from .agent_api import cell_value
    node = shutil.which("node")
    if node:
        worker = os.path.join(SKILL_DIR, "scripts", "calc", "worker.mjs")
        try:
            res = subprocess.run([node, worker], input=json.dumps({"snapshot": doc.snapshot(), "display": rects},
                                                                    ensure_ascii=False, default=str),
                                 capture_output=True, text=True, timeout=60, encoding="utf-8")
            if res.returncode == 0:
                out = json.loads(res.stdout)
                doc.set_calc(doc.version, out["values"], "node", full=True)
                return {(sid, r, c): t for sid, r, c, t in out["display"]}
        except (OSError, subprocess.TimeoutExpired, ValueError):
            pass
    texts = {}
    for sid, r1, c1, r2, c2 in rects:
        sh = doc.sheet(sid)
        for (r, c), cell in sh["cells"].items():
            if r1 <= r <= r2 and c1 <= c <= c2:
                texts[(sid, r, c)] = format_py(cell_value(doc, sh, r, c, cell), (cell.get("s") or {}).get("nf"))
    return texts


# ------------------------------------------------------------------ autofilter criteria (port of filter.js)

UI_OPS = {"contains": ("equal", "*{}*"), "notContains": ("notEqual", "*{}*"), "beginsWith": ("equal", "{}*"),
          "endsWith": ("equal", "*{}")}
EXCEL_OPS = {"equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"}
OP_ALIASES = {"=": "equal", "==": "equal", "<>": "notEqual", "!=": "notEqual", ">": "greaterThan", ">=": "greaterThanOrEqual",
              "<": "lessThan", "<=": "lessThanOrEqual"}


def to_excel_custom(op, val):
    op = OP_ALIASES.get(op, op)
    if op in UI_OPS:
        eop, pat = UI_OPS[op]
        return {"op": eop, "val": pat.format(val)}
    if op not in EXCEL_OPS:
        raise ApiError(400, f"필터 조건: {', '.join(sorted(EXCEL_OPS | set(UI_OPS)))} 또는 = <> > >= < <=")
    return {"op": op, "val": str(val)}


def _wild(pat):
    return re.compile("^" + re.escape(pat).replace(r"\*", ".*").replace(r"\?", ".") + "$", re.I)


def _cmp(a, b):
    return (a > b) - (a < b)


def custom_test(cf, value, text):
    v = str(cf.get("val", ""))
    try:
        num = float(v) if v.strip() else None
    except ValueError:
        num = None
    is_num = isinstance(value, (int, float)) and not isinstance(value, bool)
    if cf["op"].startswith(("greater", "less")) and (num is not None) != is_num:
        return False
    c = _cmp(value, num) if is_num and num is not None else _cmp(text.lower(), v.lower())
    op = cf["op"]
    if op == "equal":
        return bool(_wild(v).match(text)) if re.search(r"[*?]", v) else (c == 0 if is_num and num is not None else text.lower() == v.lower())
    if op == "notEqual":
        return not _wild(v).match(text) if re.search(r"[*?]", v) else (c != 0 if is_num and num is not None else text.lower() != v.lower())
    return {"greaterThan": c > 0, "greaterThanOrEqual": c >= 0, "lessThan": c < 0, "lessThanOrEqual": c <= 0}.get(op, True)


def filter_hidden(doc, sh, af):
    from .agent_api import cell_value
    if not af or not af.get("cols"):
        return []
    r1, c1, r2, c2 = A.parse_range(af["ref"])
    entries = [(c1 + int(k), e) for k, e in af["cols"].items()]
    texts = display_texts(doc, [(sh["id"], r1 + 1, min(c for c, _ in entries), r2, max(c for c, _ in entries))])
    hidden = []
    for r in range(r1 + 1, r2 + 1):
        for c, e in entries:
            text = texts.get((sh["id"], r, c), "")
            value = cell_value(doc, sh, r, c, sh["cells"].get((r, c)) or {})
            ok = True
            if e.get("values") is not None:
                ok = bool(e.get("blank")) if text == "" else any(x.lower() == text.lower() for x in e["values"])
            if ok and e.get("custom"):
                res = [custom_test(cf, value, text) for cf in e["custom"]]
                ok = all(res) if e.get("and") else any(res)
            if not ok:
                hidden.append(r)
                break
    return hidden


def op_filter_column(ctx, e):
    doc, sh = ctx.doc, ctx.sheet_of(e)
    af = copy.deepcopy(ctx.meta_of(sh, "autoFilter"))
    if not af:
        raise ApiError(400, "먼저 자동 필터를 켜세요: filter RANGE")
    r1, c1, r2, c2 = A.parse_range(af["ref"])
    col = A.col_index(e["column"]) if str(e["column"]).isalpha() else int(e["column"]) - 1
    if not c1 <= col <= c2:
        raise ApiError(400, f"{A.col_name(col)}열은 필터 범위({af['ref']}) 밖입니다.")
    key = str(col - c1)
    entry = {}
    if e.get("values") is not None:
        entry["values"] = [str(x) for x in e["values"]]
        if e.get("blank"):
            entry["blank"] = True
    if e.get("custom"):
        entry["custom"] = [to_excel_custom(x.get("op", "equal"), x.get("val", "")) for x in e["custom"]]
        entry["and"] = bool(e.get("and", True))
    if entry:
        af["cols"][key] = entry
    else:
        af["cols"].pop(key, None)
    ctx.meta[sh["id"]]["autoFilter"] = af
    ctx.meta[sh["id"]]["filterHidden"] = filter_hidden(doc, sh, af)
    ctx.summary.append(f"{sh['name']} {A.col_name(col)}열 필터 {'해제' if not entry else '설정'}"
                       f" (숨김 {len(ctx.meta[sh['id']]['filterHidden'])}행)")


def op_filter_clear(ctx, e):
    sh = ctx.sheet_of(e)
    af = copy.deepcopy(ctx.meta_of(sh, "autoFilter"))
    if af:
        af["cols"] = {}
        ctx.meta[sh["id"]]["autoFilter"] = af
    ctx.meta.setdefault(sh["id"], {})["filterHidden"] = []
    ctx.summary.append(f"{sh['name']} 필터 조건 모두 지우기")


def op_filter_reapply(ctx, e):
    sh = ctx.sheet_of(e)
    af = copy.deepcopy(ctx.meta_of(sh, "autoFilter"))
    if not af:
        return
    r1, c1, r2, c2 = A.parse_range(af["ref"])
    while any((r2 + 1, c) in sh["cells"] for c in range(c1, c2 + 1)):  # pick up rows typed below
        r2 += 1
    from . import region as RG
    r1, c1, r2, c2 = RG.table_body(sh, (r1, c1, r2, c2))  # but not a totals row under the data
    af["ref"] = A.range_a1(r1, c1, r2, c2)
    ctx.meta[sh["id"]]["autoFilter"] = af
    ctx.meta[sh["id"]]["filterHidden"] = filter_hidden(ctx.doc, sh, af)
    ctx.summary.append(f"{sh['name']} 필터 다시 적용")


# ------------------------------------------------------------------ charts

def op_chart_update(ctx, e):
    from .agent_api import chart_from_range, find_chart, parse_at, resolve
    sh, charts, idx = find_chart(ctx.doc, ctx.meta, ctx.meta_of, e.get("id"), e.get("sheet"), e.get("index"))
    ch = copy.deepcopy(charts[idx])
    move_to = None
    st = dict(e.get("set") or {})
    if "data" in st:
        dsh, g = resolve(ctx.doc, st.pop("data"), sh["name"])
        fresh = chart_from_range(ctx.doc, dsh, g, {"type": st.get("type", ch["type"]), "byRows": st.pop("byRows", None)},
                                 ctx.cellbuf.get(dsh["id"], {}))
        ch["series"] = fresh["series"]
    if "type" in st:
        t = st.pop("type")
        preset = {"barH": ("bar", "bar", "clustered"), "barStacked": ("bar", "col", "stacked")}.get(t)
        ch["type"], ch["barDir"], ch["grouping"] = preset or (t, "col", "standard" if t in ("line", "area") else "clustered")
        for s in ch.get("series", []):
            s.pop("type", None)
    if "at" in st:
        at_sheet, r, c = parse_at(ctx.doc, st.pop("at"))
        ch["anchor"].update(r=r, c=c, dx=0, dy=0)
        if at_sheet and at_sheet is not sh:
            move_to = at_sheet  # 'Other!E2' moves the chart to that sheet
    for k in ("width", "height"):
        if k in st:
            ch["anchor"]["w" if k == "width" else "h"] = int(st.pop(k))
    for k in ("title", "legend", "xTitle", "yTitle", "grouping", "barDir"):
        if k in st:
            ch[k] = st.pop(k)
    if "varyColors" in st:
        ch["varyColors"] = bool(st.pop("varyColors"))
    if st:
        raise ApiError(400, f"알 수 없는 차트 속성: {', '.join(st)} (type, data, title, legend, xTitle, yTitle, at, width, height, grouping, barDir, varyColors)")
    ch["dirty"] = True
    if move_to is not None:
        ch["orig"] = None  # rebuilt on save: it no longer belongs to the original sheet's drawing
        del charts[idx]
        ctx.meta_of(move_to, "charts").append(ch)
        ctx.summary.append(f"{sh['name']} 차트 {ch['id']} → '{move_to['name']}' 시트로 이동")
        return
    charts[idx] = ch
    ctx.summary.append(f"{sh['name']} 차트 {ch['id']} 수정")


# ------------------------------------------------------------------ sheets & view

def op_move_sheet(ctx, e):
    doc = ctx.doc
    sh = doc.sheet(e["sheet"])
    order = [s["id"] for s in doc.sheets if s is not sh]
    to = e.get("to")
    if isinstance(to, str) and not to.lstrip("-").isdigit():
        to = order.index(doc.sheet(to)["id"])
    to = max(0, min(int(to), len(order)))
    order.insert(to, sh["id"])
    ctx.ops.append({"op": "sheets", "order": order, "add": {}})
    ctx.structural = True
    ctx.summary.append(f"시트 '{sh['name']}' 이동 → {to + 1}번째")


def op_duplicate_sheet(ctx, e):
    from .agent_api import _check_sheet_name, _unique_name
    doc = ctx.doc
    src = doc.sheet(e["sheet"])
    name = e.get("name") or _unique_name(doc, src["name"] + " (")
    if not e.get("name"):
        i = 2
        while any(s["name"].lower() == f"{src['name']} ({i})".lower() for s in doc.sheets):
            i += 1
        name = f"{src['name']} ({i})"
    _check_sheet_name(doc, name)
    sj = doc.sheet_json(src)
    sj.update(name=name, origName=None)
    for ch in sj["charts"]:
        ch.update(id=new_id("ch"), orig=None, dirty=True)
    sid = new_id("s")
    order = [s["id"] for s in doc.sheets]
    order.insert(order.index(src["id"]) + 1, sid)
    ctx.ops.append({"op": "sheets", "order": order, "add": {sid: sj}})
    ctx.structural = True
    ctx.summary.append(f"시트 '{src['name']}' 복사 → '{name}'")


def op_sheet_prop(ctx, e):
    sh = ctx.sheet_of(e)
    kind = e["op"]
    m = ctx.meta.setdefault(sh["id"], {})
    if kind in ("hideSheet", "unhideSheet"):
        if kind == "hideSheet" and sum(1 for s in ctx.doc.sheets if s["state"] == "visible") <= 1:
            raise ApiError(400, "표시된 시트가 하나뿐이면 숨길 수 없습니다.")
        m["state"] = "hidden" if kind == "hideSheet" else "visible"
    elif kind == "tabColor":
        m["tabColor"] = e.get("color")
    elif kind == "gridlines":
        m["showGrid"] = bool(e.get("show", True))
    ctx.summary.append(f"{sh['name']} {kind}")


# ------------------------------------------------------------------ fill series (port of app.js makeSeries)

LISTS = [
    ["일", "월", "화", "수", "목", "금", "토"],
    ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"],
    ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
    ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    ["1분기", "2분기", "3분기", "4분기"], ["갑", "을", "병", "정", "무", "기", "경", "신", "임", "계"],
]


def make_series(cells):
    vals = [c.get("v") if c and not c.get("f") else None for c in cells]
    n = len(cells)
    copy_ = lambda k, base: dict(base) if base else None
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
        nf = ((cells[0] or {}).get("s") or {}).get("nf")
        if n == 1:
            return (lambda k, base: dict(base, v=vals[0] + k)) if _is_date_nf(nf) else copy_
        mx, my = (n - 1) / 2, sum(vals) / n
        sxx = sum((i - mx) ** 2 for i in range(n))
        b = sum((i - mx) * (vals[i] - my) for i in range(n)) / sxx if sxx else 0
        a = my - b * mx
        return lambda k, base: dict(base, v=float(f"{a + b * k:.15g}") if not float(f"{a + b * k:.15g}").is_integer() else int(round(a + b * k)))
    for L in LISTS:
        low = [x.lower() for x in L]
        idx = [low.index(v.lower()) if isinstance(v, str) and v.lower() in low else -1 for v in vals]
        if all(i >= 0 for i in idx):
            step = ((idx[-1] - idx[0]) / (n - 1) or 1) if n > 1 else 1
            return lambda k, base, idx=idx, L=L, step=step: dict(base, v=L[(idx[0] + round(step * k)) % len(L)])
    ms = [re.fullmatch(r"(.*?)(\d+)(\D*)", v) if isinstance(v, str) else None for v in vals]
    if all(ms) and all(m.group(1) == ms[0].group(1) and m.group(3) == ms[0].group(3) for m in ms):
        nums = [int(m.group(2)) for m in ms]
        step = ((nums[-1] - nums[0]) / (n - 1) or 1) if n > 1 else 1
        width = len(ms[0].group(2))
        pad = ms[0].group(2).startswith("0")

        def gen(k, base):
            v = round(nums[0] + step * k)
            return dict(base, v=ms[0].group(1) + (str(abs(v)).zfill(width) if pad else str(v)) + ms[0].group(3))
        return gen
    return copy_


def op_fill_series(ctx, e):
    from .agent_api import resolve
    sh, (sr1, sc1, sr2, sc2) = resolve(ctx.doc, e["source"], e.get("sheet"))
    _, (tr1, tc1, tr2, tc2) = resolve(ctx.doc, e["target"], sh["name"])
    vertical = (tc1, tc2) == (sc1, sc2) and (tr1 != sr1 or tr2 != sr2)
    if not vertical and (tr1, tr2) != (sr1, sr2):
        raise ApiError(400, "target은 source를 위/아래 또는 왼쪽/오른쪽으로 늘린 범위여야 합니다(예: source A2:A3, target A2:A20).")
    lines = range(sc1, sc2 + 1) if vertical else range(sr1, sr2 + 1)
    hidden = set(sh["filterHidden"])
    for line in lines:
        span = range(sr1, sr2 + 1) if vertical else range(sc1, sc2 + 1)
        seq = [ctx.cur(sh, i, line) if vertical else ctx.cur(sh, line, i) for i in span]
        start, length = (sr1 if vertical else sc1), len(seq)
        gen = make_series(seq)
        for p in range(tr1 if vertical else tc1, (tr2 if vertical else tc2) + 1):
            if start <= p < start + length or (vertical and p in hidden):
                continue
            k = p - start
            i = k % length
            base = seq[i]
            if base and base.get("f"):
                cell = dict(base, f=refs.translate(base["f"], k - i if vertical else 0, 0 if vertical else k - i))
            else:
                cell = gen(k, base or {}) if base else None
            if cell:
                cell.pop("a", None)
            ctx.put(sh, p, line, cell) if vertical else ctx.put(sh, line, p, cell)
    ctx.summary.append(f"{sh['name']}!{A.range_a1(tr1, tc1, tr2, tc2)} 연속 데이터 채우기")


# ------------------------------------------------------------------ copy / move / shift cells / duplicates

def op_copy(ctx, e, move=False):
    from .agent_api import cell_value, resolve
    ssh, (r1, c1, r2, c2) = resolve(ctx.doc, e["from"], e.get("sheet"))
    tsh, (tr, tc, _, _) = resolve(ctx.doc, e["to"], e.get("sheet") or ssh["name"])
    what = "all" if move else e.get("what", "all")
    transpose = bool(e.get("transpose")) and not move
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
        raise ApiError(400, "한 번에 복사할 수 있는 셀은 200,000개까지입니다.")
    block = [[(ctx.cur(ssh, r, c), cell_value(ctx.doc, ssh, r, c, ctx.cur(ssh, r, c) or {})) for c in range(c1, c2 + 1)]
             for r in range(r1, r2 + 1)]
    if move:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ctx.put(ssh, r, c, None)
    for i, row in enumerate(block):
        for j, (src, val) in enumerate(row):
            di, dj = (j, i) if transpose else (i, j)
            r, c = tr + di, tc + dj
            dst = ctx.cur(tsh, r, c)
            if what == "all":
                cell = dict(src) if src else None
                if cell and cell.get("f") and not move:
                    cell["f"] = refs.translate(cell["f"], r - (r1 + i), c - (c1 + j))
            elif what == "values":
                cell = {k: v for k, v in (dst or {}).items() if k in ("s", "h", "cm")}
                if val is not None:
                    cell["v"] = val
            elif what == "formats":
                cell = {k: v for k, v in (dst or {}).items() if k != "s"}
                if src and src.get("s"):
                    cell["s"] = src["s"]
            elif what == "formulas":
                cell = {k: v for k, v in (src or {}).items() if k != "s"}
                if (dst or {}).get("s"):
                    cell["s"] = dst["s"]
                if cell.get("f"):
                    cell["f"] = refs.translate(cell["f"], r - (r1 + i), c - (c1 + j))
            else:
                raise ApiError(400, "what: all | values | formats | formulas")
            if cell:
                cell.pop("a", None)
            ctx.put(tsh, r, c, cell or None)
    h, w = (c2 - c1 + 1, r2 - r1 + 1) if transpose else (r2 - r1 + 1, c2 - c1 + 1)
    ctx.summary.append(f"{ssh['name']}!{A.range_a1(r1, c1, r2, c2)} {'이동' if move else '복사'} → "
                       f"{tsh['name']}!{A.range_a1(tr, tc, tr + h - 1, tc + w - 1)}{' (행/열 바꿈)' if transpose else ''}")


def op_shift_cells(ctx, e):
    from .agent_api import resolve
    sh, (r1, c1, r2, c2) = resolve(ctx.doc, e["range"], e.get("sheet"))
    d = e.get("shift") or ("right" if e["op"] == "insertCells" else "left")
    valid = ("right", "down") if e["op"] == "insertCells" else ("left", "up")
    if d not in valid:
        raise ApiError(400, f"shift: {' | '.join(valid)}")
    h, w = r2 - r1 + 1, c2 - c1 + 1
    keys = set(sh["cells"]) | set(ctx.cellbuf.get(sh["id"], {}))
    moves = []
    for (r, c) in keys:
        cell = ctx.cur(sh, r, c)
        if not cell:
            continue
        if d in ("right", "left") and r1 <= r <= r2 and c >= c1:
            moves.append((r, c, cell))
        if d in ("down", "up") and c1 <= c <= c2 and r >= r1:
            moves.append((r, c, cell))
    for r, c, _ in moves:
        ctx.put(sh, r, c, None)
    for r, c, cell in moves:
        if d == "right":
            ctx.put(sh, r, c + w, cell)
        elif d == "down":
            ctx.put(sh, r + h, c, cell)
        elif d == "left" and c > c2:
            ctx.put(sh, r, c - w, cell)
        elif d == "up" and r > r2:
            ctx.put(sh, r - h, c, cell)
    ctx.structural, ctx.transformable = True, False
    ctx.summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 셀 {'삽입' if e['op'] == 'insertCells' else '삭제'}({d})")


def op_remove_duplicates(ctx, e):
    from .agent_api import cell_value, clamp, resolve
    from . import region as RG
    sh, g = resolve(ctx.doc, e["range"], e.get("sheet"))
    g, _ = RG.auto_range(sh, g, ctx.cur)
    r1, c1, r2, c2 = clamp(sh, g, pending=ctx.cellbuf.get(sh["id"]))
    header = e.get("header")
    if header is None:
        header = any(isinstance(cell_value(ctx.doc, sh, r1, c, ctx.cur(sh, r1, c) or {}), str) for c in range(c1, c2 + 1))
    cols = [A.col_index(x) for x in e.get("columns") or []] or list(range(c1, c2 + 1))
    start = r1 + (1 if header else 0)
    seen, keep = set(), []
    for r in range(start, r2 + 1):
        k = tuple(json.dumps(v.lower() if isinstance(v, str) else v, default=str)
                  for v in (cell_value(ctx.doc, sh, r, c, ctx.cur(sh, r, c) or {}) for c in cols))
        if k not in seen:
            seen.add(k)
            keep.append(r)
    rows = [[ctx.cur(sh, r, c) for c in range(c1, c2 + 1)] for r in keep]
    for r in range(start, r2 + 1):
        for c in range(c1, c2 + 1):
            ctx.put(sh, r, c, None)
    for i, row in enumerate(rows):
        dst = start + i
        for j, cell in enumerate(row):
            if cell and cell.get("f") and dst != keep[i]:
                cell = dict(cell, f=refs.translate(cell["f"], dst - keep[i], 0))
            ctx.put(sh, dst, c1 + j, cell)
    ctx.structural, ctx.transformable = True, False
    ctx.summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 중복 {(r2 - start + 1) - len(keep)}행 제거")


# ------------------------------------------------------------------ find / replace, comments, links

def _matcher(text, case, whole):
    body = re.escape(text).replace(r"\*", ".*").replace(r"\?", ".").replace(r"\~\.\*", r"\*")
    return re.compile(("^" if whole else "") + body + ("$" if whole else ""), 0 if case else re.I)


def op_replace(ctx, e):
    rx = _matcher(e["find"], e.get("matchCase"), e.get("whole"))
    rep = str(e.get("with", ""))
    where = e.get("in", "values")
    sheets = ctx.doc.sheets if e.get("all") else [ctx.sheet_of(e)]
    n = 0
    for sh in sheets:
        for (r, c), cell in list(sh["cells"].items()):
            cell = ctx.cur(sh, r, c) or cell
            if where == "formulas" and cell.get("f"):
                nf = rx.sub(lambda m: rep, cell["f"])
                if nf != cell["f"]:
                    ctx.put(sh, r, c, dict(cell, f=nf))
                    n += 1
            elif not cell.get("f") and isinstance(cell.get("v"), str):
                nv = rx.sub(lambda m: rep, cell["v"])
                if nv != cell["v"]:
                    ctx.put(sh, r, c, dict(cell, v=nv))
                    n += 1
    ctx.summary.append(f"바꾸기 '{e['find']}' → '{rep}' {n}개")


def op_comment(ctx, e):
    from .agent_api import resolve
    sh, (r, c, _, _) = resolve(ctx.doc, e["cell"], e.get("sheet"))
    cell = dict(ctx.cur(sh, r, c) or {})
    if e.get("text"):
        cell["cm"] = str(e["text"])
    else:
        cell.pop("cm", None)
    ctx.put(sh, r, c, cell or None)
    ctx.summary.append(f"{sh['name']}!{A.a1(r, c)} 메모 {'설정' if e.get('text') else '삭제'}")


def op_link(ctx, e):
    from .agent_api import resolve
    sh, (r, c, _, _) = resolve(ctx.doc, e["cell"], e.get("sheet"))
    cell = dict(ctx.cur(sh, r, c) or {})
    url = e.get("url")
    if url:
        cell["h"] = url
        if e.get("text") is not None:
            cell["v"] = str(e["text"])
        elif cell.get("v") is None and not cell.get("f"):
            cell["v"] = url
        st = dict(cell.get("s") or {})
        st.setdefault("fc", "#0563C1")
        st.setdefault("u", True)
        cell["s"] = st
    else:
        cell.pop("h", None)
    ctx.put(sh, r, c, cell or None)
    ctx.summary.append(f"{sh['name']}!{A.a1(r, c)} 하이퍼링크 {'설정' if url else '삭제'}")


# ------------------------------------------------------------------ autofit, format copy, decimals

def _text_width(s):
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in s)


def op_autofit(ctx, e):
    from .agent_api import resolve
    sh, (r1, c1, r2, c2) = resolve(ctx.doc, e["range"], e.get("sheet"))
    is_cols = r1 == 0 and r2 == A.MAX_ROWS - 1
    lo, hi = (c1, c2) if is_cols else (r1, r2)
    if hi - lo > 2000:
        raise ApiError(400, "자동 맞춤은 한 번에 2,000개 행/열까지입니다.")
    keys = [(r, c) for (r, c) in set(sh["cells"]) | set(ctx.cellbuf.get(sh["id"], {})) if lo <= (c if is_cols else r) <= hi]
    texts = display_texts(ctx.doc, [(sh["id"], min((k[0] for k in keys), default=0), min((k[1] for k in keys), default=0),
                                     max((k[0] for k in keys), default=0), max((k[1] for k in keys), default=0))]) if keys else {}
    dims = ctx.meta_of(sh, "cols" if is_cols else "rows")
    best = {}
    for (r, c) in keys:
        cell = ctx.cur(sh, r, c) or {}
        st = cell.get("s") or {}
        text = texts.get((sh["id"], r, c)) if (r, c) in sh["cells"] else str(cell.get("v", ""))
        if not text:
            continue
        scale = float(st.get("fs", ctx.doc.default_font.get("size", 11))) / 11
        i = c if is_cols else r
        if is_cols:
            best[i] = max(best.get(i, 0), max(_text_width(line) for line in str(text).split("\n")) * scale)
        else:
            best[i] = max(best.get(i, 0), (str(text).count("\n") + 1) * scale)
    for i in range(lo, hi + 1):
        d = dict(dims.get(str(i)) or {})
        if i in best:
            if is_cols:
                d["w"] = max(sh["defaultColW"], round(best[i] * 7 + 12))  # ~7px per character (Calibri 11)
            else:
                d["h"] = max(sh["defaultRowH"], round(best[i] * 11 * 1.25 * 96 / 72 + 5))
        else:
            d.pop("w" if is_cols else "h", None)
        if d:
            dims[str(i)] = d
        else:
            dims.pop(str(i), None)
    ctx.summary.append(f"{sh['name']} {'열 너비' if is_cols else '행 높이'} 자동 맞춤")


def op_copy_format(ctx, e):
    from .agent_api import clamp, resolve
    ssh, (r1, c1, r2, c2) = resolve(ctx.doc, e["from"], e.get("sheet"))
    tsh, g = resolve(ctx.doc, e["to"], e.get("sheet") or ssh["name"])
    tr1, tc1, tr2, tc2 = clamp(tsh, g, pending=ctx.cellbuf.get(tsh["id"]))
    h, w = r2 - r1 + 1, c2 - c1 + 1
    for r in range(tr1, tr2 + 1):
        for c in range(tc1, tc2 + 1):
            src = ctx.cur(ssh, r1 + (r - tr1) % h, c1 + (c - tc1) % w) or {}
            cell = {k: v for k, v in (ctx.cur(tsh, r, c) or {}).items() if k != "s"}
            if src.get("s"):
                cell["s"] = src["s"]
            ctx.put(tsh, r, c, cell or None)
    ctx.summary.append(f"서식 복사 {ssh['name']}!{A.range_a1(r1, c1, r2, c2)} → {tsh['name']}!{A.range_a1(tr1, tc1, tr2, tc2)}")


def op_decimals(ctx, e):
    from .agent_api import cell_value, clamp, resolve
    sh, g = resolve(ctx.doc, e["range"], e.get("sheet"))
    r1, c1, r2, c2 = clamp(sh, g, pending=ctx.cellbuf.get(sh["id"]))
    delta = int(e.get("delta", 1))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = dict(ctx.cur(sh, r, c) or {})
            st = dict(cell.get("s") or {})
            nf = st.get("nf") if st.get("nf") not in (None, "General") else None
            if not nf:
                v = cell_value(ctx.doc, sh, r, c, cell)
                dec = len(f"{v:.10g}".split(".")[1]) if isinstance(v, float) and "." in f"{v:.10g}" else 0
                nf = "0." + "0" * dec if dec else "0"
            parts = []
            for p in nf.split(";"):
                for _ in range(abs(delta)):
                    if delta > 0:
                        p = re.sub(r"0\.(0*)", lambda m: "0." + m.group(1) + "0", p, count=1) if re.search(r"0\.0*", p) else re.sub(r"0(?![0.])", "0.0", p, count=1)
                    else:
                        p = re.sub(r"0\.(0*)", lambda m: "0." + m.group(1)[1:] if len(m.group(1)) > 1 else "0", p, count=1)
                parts.append(p)
            st["nf"] = ";".join(parts)
            cell["s"] = st
            ctx.put(sh, r, c, cell)
    ctx.summary.append(f"{sh['name']}!{A.range_a1(r1, c1, r2, c2)} 소수 자릿수 {'+' if delta > 0 else ''}{delta}")


HANDLERS = {
    "filterColumn": op_filter_column, "filterClear": op_filter_clear, "filterReapply": op_filter_reapply,
    "chartUpdate": op_chart_update, "moveSheet": op_move_sheet, "duplicateSheet": op_duplicate_sheet,
    "hideSheet": op_sheet_prop, "unhideSheet": op_sheet_prop, "tabColor": op_sheet_prop, "gridlines": op_sheet_prop,
    "fillSeries": op_fill_series, "copy": op_copy, "move": lambda ctx, e: op_copy(ctx, e, move=True),
    "insertCells": op_shift_cells, "deleteCells": op_shift_cells, "removeDuplicates": op_remove_duplicates,
    "replace": op_replace, "comment": op_comment, "link": op_link, "autofit": op_autofit,
    "copyFormat": op_copy_format, "decimals": op_decimals,
}
