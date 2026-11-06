"""Data-region detection shared by agent commands (mirrors assets/web/js/filter.js).

Tables rarely start at A1: there may be a title, blank rows, notes beside them and a totals row
right under the data. These helpers find the table around a cell and the "body" to filter/sort.
"""
from __future__ import annotations

import re

from . import a1 as A
from . import refs

# Total-row labels in the data, English and Korean (합계 = total, 소계 = subtotal, 평균 = average, ...).
TOTAL_LABEL = re.compile(r"^\s*(합\s*계|총\s*계|소\s*계|계|누\s*계|평\s*균|total|totals|grand total|sum|subtotal|average)\s*$", re.I)
AGG_FN = re.compile(r"\b(SUM|SUBTOTAL|AVERAGE|COUNT|COUNTA|MAX|MIN|AGGREGATE)\s*\(", re.I)


def _filled(sh, r, c, cur=None):
    cell = (cur(sh, r, c) if cur else sh["cells"].get((r, c))) or {}
    return bool(cell.get("f")) or (cell.get("v") not in (None, ""))


def current_region(sh, r, c, cur=None):
    """Excel's current region (Ctrl+A / Ctrl+Shift+L): expand while a neighbouring row/col has data."""
    r1 = r2 = r
    c1 = c2 = c

    def row_has(rr):
        return rr >= 0 and any(_filled(sh, rr, cc, cur) for cc in range(max(0, c1 - 1), c2 + 2))

    def col_has(cc):
        return 0 <= cc < A.MAX_COLS and any(_filled(sh, rr, cc, cur) for rr in range(max(0, r1 - 1), r2 + 2))

    changed = True
    while changed:
        changed = False
        while r1 > 0 and row_has(r1 - 1):
            r1 -= 1; changed = True
        while r2 < A.MAX_ROWS - 1 and row_has(r2 + 1):
            r2 += 1; changed = True
        while c1 > 0 and col_has(c1 - 1):
            c1 -= 1; changed = True
        while c2 < A.MAX_COLS - 1 and col_has(c2 + 1):
            c2 += 1; changed = True
    return r1, c1, r2, c2


def is_total_row(sh, r, r1, c1, c2, cur=None):
    """A row under the data that summarises it: a total label (TOTAL_LABEL) or an aggregate over the rows above."""
    get = (lambda rr, cc: (cur(sh, rr, cc) if cur else sh["cells"].get((rr, cc))) or {})
    first = get(r, c1).get("v")
    if isinstance(first, str) and TOTAL_LABEL.match(first):
        return True
    for c in range(c1, c2 + 1):
        f = get(r, c).get("f")
        if f and AGG_FN.search(f):
            for _, _, ref in refs.ref_tokens(f):
                b = ref["b"] or ref["a"]
                lo, hi = min(ref["a"]["r"], b["r"]), max(ref["a"]["r"], b["r"])
                if ref["kind"] == "cell" and hi < r and hi - lo >= 1 and lo >= r1:
                    return True
    return False


def table_body(sh, g, cur=None):
    """Trim trailing total rows from a region: (r1, c1, r2, c2) of header + data."""
    r1, c1, r2, c2 = g
    while r2 > r1 + 1 and is_total_row(sh, r2, r1, c1, c2, cur):
        r2 -= 1
    return r1, c1, r2, c2


def first_filled(sh, g, cur=None):
    r1, c1, r2, c2 = g
    best = None
    for (r, c) in sh["cells"]:
        if r1 <= r <= r2 and c1 <= c <= c2 and _filled(sh, r, c, cur):
            if best is None or (r, c) < best:
                best = (r, c)
    return best


def auto_range(sh, g, cur=None, *, trim_totals=True):
    """Range an agent meant when it gave one cell or whole rows/columns: the table there."""
    r1, c1, r2, c2 = g
    single = (r1, c1) == (r2, c2)
    whole = (r1 == 0 and r2 == A.MAX_ROWS - 1) or (c1 == 0 and c2 == A.MAX_COLS - 1)
    if not (single or whole):
        return g, False
    start = (r1, c1) if single and _filled(sh, r1, c1, cur) else first_filled(sh, g, cur)
    if not start:
        return g, False
    reg = current_region(sh, *start, cur)
    if whole:  # stay inside the requested columns/rows
        reg = (max(reg[0], r1), max(reg[1], c1), min(reg[2], r2), min(reg[3], c2))
    return (table_body(sh, reg, cur) if trim_totals else reg), True


def tables(sh, limit=8):
    """All blocks of contiguous data on a sheet (largest first): [(r1, c1, r2, c2)]."""
    seen, out = set(), []
    for (r, c) in sorted(sh["cells"]):
        if (r, c) in seen or not _filled(sh, r, c):
            continue
        g = current_region(sh, r, c)
        for (rr, cc) in sh["cells"]:
            if g[0] <= rr <= g[2] and g[1] <= cc <= g[3]:
                seen.add((rr, cc))
        out.append(g)
    out.sort(key=lambda g: -((g[2] - g[0] + 1) * (g[3] - g[1] + 1)))
    return out[:limit]
