"""Pivot tables (step 1: safeguards). Sheet meta "pivots" (from xlsx_io): each pivot's output range
`ref`, its `source` ('Sheet'!A1:D13), fields, and `idx` in the file. Shared by the edit guard
(session), structural edits (agent_api) and describe."""
from __future__ import annotations

from . import a1 as A


def ranges(sh):
    """[(pivot, (r1, c1, r2, c2))] of the pivots on a sheet."""
    out = []
    for p in sh.get("pivots") or []:
        try:
            out.append((p, A.parse_range(p["ref"])))
        except (KeyError, ValueError):
            pass
    return out


def at(sh, r, c):
    """The pivot whose output covers cell (r, c), or None."""
    for p, (r1, c1, r2, c2) in ranges(sh):
        if r1 <= r <= r2 and c1 <= c <= c2:
            return p
    return None


def overlapping(sh, g):
    """The pivots whose output intersects range g = (r1, c1, r2, c2)."""
    return [p for p, (r1, c1, r2, c2) in ranges(sh) if r1 <= g[2] and g[0] <= r2 and c1 <= g[3] and g[1] <= c2]


def blocked_message(p):
    return (f"The area of pivot table '{p['name']}' ({p['ref']}) cannot be changed. Its results come from the source data "
            f"({p.get('source') or '?'}), so change the source (after saving, Excel refreshes the pivot when it opens the file).")


def shift_problem(sh, axis, at_, count):
    """Why inserting/deleting rows or columns here would break a pivot (Excel refuses too), else None.
    Inserting at the pivot's first row/column or before it moves the pivot; inside it, or deleting any
    of its rows/columns, would cut through it."""
    for p, (r1, c1, r2, c2) in ranges(sh):
        lo, hi = (r1, r2) if axis == "row" else (c1, c2)
        what = "rows" if axis == "row" else "columns"
        if count > 0 and lo < at_ <= hi:
            return f"Cannot insert {what} inside pivot table '{p['name']}' ({p['ref']}). Insert them above, below, left or right of it."
        if count < 0 and at_ <= hi and lo <= at_ - count - 1:
            return f"Cannot delete {what} of pivot table '{p['name']}' ({p['ref']})."
    return None


# ------------------------------------------------------------------ step 2: refresh in place
import math
import statistics as _st


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, float) and math.isnan(v))


def aggregate(fn, vals):
    """Excel's pivot summary functions over the field's values (blank -> None)."""
    nums = [v for v in vals if _num(v)]
    fn = (fn or "sum").lower()
    if fn == "count":  # count of non-empty values (Excel: COUNTA)
        n = sum(1 for v in vals if v is not None and v != "")
        return n if n else None
    if fn == "countnums":
        return len(nums) or None
    if not nums:
        return None
    if fn == "sum":
        return sum(nums)
    if fn == "average":
        return sum(nums) / len(nums)
    if fn == "max":
        return max(nums)
    if fn == "min":
        return min(nums)
    if fn == "product":
        return math.prod(nums)
    if fn in ("stddev", "var"):
        if len(nums) < 2:
            return None
        return _st.stdev(nums) if fn == "stddev" else _st.variance(nums)
    if fn in ("stddevp", "varp"):
        return _st.pstdev(nums) if fn == "stddevp" else _st.pvariance(nums)
    return sum(nums)


def _same(a, b):
    if _num(a) and _num(b):
        return abs(float(a) - float(b)) < 1e-9
    return (str(a).strip().lower() if a is not None else "") == (str(b).strip().lower() if b is not None else "")


def _lines(items, fields, axis_items, labels):
    """Each output row (or column) of the data area as (kind, {field: value}, data index):
    kind 'cell' (a combination, subtotal or grand total — the filter holds the fields it fixes) or
    'blank'. From Excel's rowItems/colItems when the file has them; otherwise (LibreOffice writes
    none) from the label cells, for at most one field on the axis."""
    pos_fields = fields
    if axis_items is not None:
        out, prev = [], []
        for it in axis_items:
            if it["t"] == "blank":
                out.append(("blank", {}, 0))
                continue
            tup = (prev[: it["r"]] + it["x"]) if it["t"] != "grand" else []
            prev = tup if it["t"] != "grand" else prev
            flt = {pos_fields[k]: v for k, v in enumerate(tup) if k < len(pos_fields) and pos_fields[k] != "Σ" and v is not None}
            out.append(("cell", flt, it.get("d", it.get("i", 0))))
        return out, None
    real = [f for f in fields if f != "Σ"]
    if len(real) > 1:
        return None, "a multi-field pivot from a file without layout records (rowItems/colItems)"
    out = []
    for n, lab in enumerate(labels):
        if not real:
            out.append(("cell", {}, 0))
        elif any(_same(lab, v) for v in items.get(real[0], [])):
            out.append(("cell", {real[0]: lab}, 0))
        elif n == len(labels) - 1:
            out.append(("cell", {}, 0))  # the last line with an unknown label: grand total
        else:
            out.append(("blank", {}, 0))
    return out, None


def refresh(doc, sh, p, value_at):
    """Recompute a pivot's data cells from its source, in the layout the file has.
    value_at(sheet_dict, r, c) -> the cell's current value. Returns (cells, notes): cells as
    [r, c, cell or None] for the data cells that change; notes in Korean for the user/agent."""
    lay = p.get("layout") or {}
    notes = []
    src = p.get("source") or ""
    sheet_name, ref = A.split_sheet(src) if "!" in src else (None, None)
    ssh = next((s for s in doc.sheets if sheet_name and s["name"].lower() == sheet_name.lower()), None)
    if ssh is None or not ref:
        return [], [f"'{p['name']}': source ({src or 'none'}) not found; not refreshed."]
    r1, c1, r2, c2 = A.parse_range(ref)
    header = [value_at(ssh, r1, c) for c in range(c1, c2 + 1)]
    fields = lay.get("fields") or []
    col_of = {}
    for k, name in enumerate(fields):
        hit = next((j for j, h in enumerate(header) if _same(h, name)), None)
        col_of[name] = hit if hit is not None else (k if k < len(header) else None)
    records = []
    for r in range(r1 + 1, r2 + 1):
        rec = {name: (value_at(ssh, r, c1 + j) if j is not None else None) for name, j in col_of.items()}
        if any(v is not None and v != "" for v in rec.values()):
            records.append(rec)
    for f, hid in (lay.get("hidden") or {}).items():
        records = [rec for rec in records if not any(_same(rec.get(f), h) for h in hid)]

    o1, oc1, o2, oc2 = A.parse_range(p["ref"])
    dr, dc = o1 + (lay.get("firstDataRow") or 0), oc1 + (lay.get("firstDataCol") or 0)
    if dr > o2 or dc > oc2:
        return [], [f"'{p['name']}': no data area."]
    items = lay.get("items") or {}
    row_labels = [value_at(sh, r, dc - 1) if dc > oc1 else None for r in range(dr, o2 + 1)]
    col_labels = [value_at(sh, dr - 1, c) if dr > o1 else None for c in range(dc, oc2 + 1)]
    rows, why_r = _lines(items, lay.get("rowFields") or [], lay.get("rowItems"), row_labels)
    cols, why_c = _lines(items, lay.get("colFields") or [], lay.get("colItems"), col_labels)
    values = p.get("values") or []
    if rows is None or cols is None or (lay.get("rowItems") is None and len(values) > 1):
        why = why_r or why_c or "several value fields in a file without layout records"
        return [], [f"'{p['name']}': cannot be refreshed here ({why}). After saving, Excel refreshes it when it opens the file."]
    if len(rows) != o2 - dr + 1 or len(cols) != oc2 - dc + 1:
        notes.append(f"'{p['name']}': the file's layout ({len(rows)} rows × {len(cols)} columns) differs from the area size; only the overlap was computed.")
    out = []
    for i, (rk, rf, rd) in enumerate(rows[: o2 - dr + 1]):
        for j, (ck, cf, cd) in enumerate(cols[: oc2 - dc + 1]):
            if rk == "blank" or ck == "blank":
                continue
            di = rd if "Σ" in (lay.get("rowFields") or []) else cd if "Σ" in (lay.get("colFields") or []) else 0
            if di >= len(values):
                continue
            flt = {**rf, **cf}
            vals = [rec.get(values[di]["field"]) for rec in records if all(_same(rec.get(f), v) for f, v in flt.items())]
            new = aggregate(values[di]["fn"], vals)
            r, c = dr + i, dc + j
            old = sh["cells"].get((r, c)) or {}
            if not _same(old.get("v"), new) or (new is None) != (old.get("v") is None):
                cell = {k: v for k, v in old.items() if k not in ("v", "f", "t")}
                if new is not None:
                    cell["v"] = new
                out.append([r, c, cell or None])
    axis_fields = [f for f in (lay.get("rowFields") or []) + (lay.get("colFields") or []) if f != "Σ"]
    new_items = {f: sorted({str(rec.get(f)) for rec in records if rec.get(f) not in (None, "") and not any(_same(rec.get(f), v) for v in items.get(f, []))}) for f in axis_fields}
    new_items = {f: v for f, v in new_items.items() if v}
    if new_items:
        notes.append(f"'{p['name']}': the source has new items ({'; '.join(f'{f}: ' + ', '.join(v[:5]) for f, v in new_items.items())}). "
                     "Their rows and columns appear when Excel refreshes the pivot (on opening, after saving). They are included in the grand total.")
    return out, notes


# ------------------------------------------------------------------ step 3: create
# The words Excel writes into a pivot table come from Excel's UI language. A new pivot uses one set
# ("en" by default; "ko" = Korean Excel's wording, used when the Korean preview creates the pivot).
LABELS = {
    "en": {"fn": {"sum": "Sum", "count": "Count", "average": "Average", "max": "Max", "min": "Min", "product": "Product",
                  "countNums": "Count Numbers", "stdDev": "StdDev", "stdDevp": "StdDevp", "var": "Var", "varp": "Varp"},
           "value": "{fn} of {field}", "subtotal": "{item} Total", "grand": "Grand Total", "all": "(All)", "values": "Values",
           "sheet": "Pivot", "table": "PivotTable{n}"},
    "ko": {"fn": {"sum": "합계", "count": "개수", "average": "평균", "max": "최대", "min": "최소", "product": "곱",
                  "countNums": "숫자 개수", "stdDev": "표준 편차", "stdDevp": "표준 편차(전체)", "var": "분산", "varp": "분산(전체)"},
           "value": "{fn} : {field}", "subtotal": "{item} 요약", "grand": "총합계", "all": "(모두)", "values": "값",
           "sheet": "피벗", "table": "피벗 테이블{n}"},
}
FUNCS = list(LABELS["en"]["fn"])
FN_NAMES = {k.lower(): k for k in FUNCS}


def labels(lang):
    return LABELS.get(lang or "en") or LABELS["en"]


def is_total_label(x):
    """A row label of a subtotal or grand total line, in any label language."""
    return isinstance(x, str) and any(x == L["grand"] or x.endswith(L["subtotal"].replace("{item}", "")) for L in LABELS.values())


def _sort_key(v):
    return (0, float(v), "") if _num(v) else (1, 0.0, str(v))


def build(source_rows, header, rows, cols, values, filters, lang="en"):
    """Lay out a new pivot table in tabular form (Excel: Show in Tabular Form, subtotals at the bottom
    of each outer group, grand totals). source_rows: list of dicts by header name. Returns
    {"grid": [[value, ...], ...] (the whole output, filters on top), "table_top": rows above the
    table, "layout": {...like xlsx_io._pivot_layout}, "items": {field: sorted items}, "values": [...]}.
    Raises ValueError with a message for unsupported or invalid specs. lang picks the labels (LABELS)."""
    L = labels(lang)
    names = list(header)
    for f in rows + cols + filters + [v["field"] for v in values]:
        if f not in names:
            raise ValueError(f"The source has no column '{f}'. Headers: {', '.join(map(str, names))}")
    if not rows:
        raise ValueError("At least one row field (--rows) is needed.")
    if not values:
        raise ValueError("At least one value field (--values) is needed.")
    if len(cols) > 1:
        raise ValueError("At most one column field is supported.")
    if cols and len(values) > 1:
        raise ValueError("With a column field, at most one value field is supported.")
    used = rows + cols + filters
    if len(set(used)) != len(used):
        raise ValueError("A field can be used only once across rows, columns and filters.")
    for v in values:
        fn = FN_NAMES.get((v.get("fn") or "sum").lower())
        if fn is None:
            raise ValueError(f"Unsupported summary function '{v.get('fn')}'. Supported: {', '.join(FUNCS)}")
        v["fn"] = fn
        v.setdefault("name", L["value"].format(fn=L["fn"][fn], field=v["field"]))
    items = {f: sorted({r.get(f) for r in source_rows if r.get(f) not in (None, "")}, key=_sort_key) for f in rows + cols + filters}

    def agg(v, flt):
        vals = [r.get(v["field"]) for r in source_rows if all(_same(r.get(f), x) for f, x in flt.items())]
        return aggregate(v["fn"], vals)

    # rows (tabular form): every combination present, a subtotal row after each outer group, grand total
    def gen(depth, prefix, flt):
        f = rows[depth]
        for x in items[f]:
            here = {**flt, f: x}
            if not any(all(_same(r.get(k), y) for k, y in here.items()) for r in source_rows):
                continue
            if depth == len(rows) - 1:
                yield "data", prefix + [x], here
            else:
                yield from gen(depth + 1, prefix + [x], here)
                yield "sub", prefix + [x], here

    def common(a, b):
        n = 0
        while n < min(len(a), len(b)) and _same(a[n], b[n]):
            n += 1
        return n

    row_lines, prev, prev_data = [], [], []
    for kind, tup, flt in gen(0, [], {}):
        r = common(prev, tup) if kind == "data" else len(tup) - 1
        r = min(r, len(tup) - 1)
        if kind == "data":
            shown = common(prev_data, tup)
            line_labels = [None] * shown + tup[shown:]
            prev_data = tup
        else:
            line_labels = [None] * len(rows)
            line_labels[len(tup) - 1] = L["subtotal"].format(item=tup[-1])
        row_lines.append((line_labels, flt, {"t": "data" if kind == "data" else "default", "r": r, "i": 0, "x": tup[r:]}))
        prev = tup
    row_lines.append(([L["grand"]] + [None] * (len(rows) - 1), {}, {"t": "grand", "r": 0, "i": 0, "x": [None]}))

    if cols:
        col_lines = [({cols[0]: x}, {"t": "data", "r": 0, "i": 0, "x": [x]}, x) for x in items[cols[0]]] + [({}, {"t": "grand", "r": 0, "i": 0, "x": [None]}, L["grand"])]
    elif len(values) > 1:
        col_lines = [({}, {"t": "data", "r": 0, "i": k, "x": [None], "d": k}, v["name"]) for k, v in enumerate(values)]
    else:
        col_lines = [({}, {"t": "data", "r": 0, "i": 0, "x": [None]}, values[0]["name"])]

    grid = []
    for f in filters:
        grid.append([f, L["all"]] + [None] * max(0, len(rows) + len(col_lines) - 2))
    if filters:
        grid.append([None] * (len(rows) + len(col_lines)))
    top = len(grid)
    width = len(rows) + len(col_lines)
    sigma = not cols and len(values) > 1  # several data fields across the columns: a "Values" header row (Excel, LibreOffice)
    if cols:
        grid.append([values[0]["name"]] + [None] * (len(rows) - 1) + [cols[0]] + [None] * (len(col_lines) - 1))
    elif sigma:
        grid.append([None] * len(rows) + [L["values"]] + [None] * (len(col_lines) - 1))
    grid.append(list(rows) + [lab for _, _, lab in col_lines])
    first_data_row = len(grid) - top
    for row_labels, rflt, _ in row_lines:
        line = list(row_labels)
        for k, (cflt, it, _) in enumerate(col_lines):
            v = values[it.get("d", 0)]
            line.append(agg(v, {**rflt, **cflt}))
        grid.append(line)
    layout = {"rowFields": list(rows), "colFields": (list(cols) if cols else (["Σ"] if len(values) > 1 else [])),
              "rowItems": [it for _, _, it in row_lines], "colItems": [it for _, it, _ in col_lines],
              "hidden": {}, "firstHeaderRow": 1 if (cols or sigma) else 0, "firstDataRow": first_data_row, "firstDataCol": len(rows),
              "items": {f: v for f, v in items.items()}, "fields": names}
    return {"grid": grid, "table_top": top, "width": width, "layout": layout, "items": items, "values": values}
