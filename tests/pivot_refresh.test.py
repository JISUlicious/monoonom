"""Pivot refresh engine (server/pivots.refresh): python3 tests/pivot_refresh.test.py

Synthetic pivots in the shape xlsx_io reads from files. The Excel-style layout (rowItems / colItems)
follows ECMA-376 Part 1 §18.10.1.84 (rowItems) / §18.10.1.17 (colItems): each <i> lists what one
output row/column shows; r = leading fields repeated from the previous <i>, x = item per field,
t = data | default (subtotal) | grand | blank, i / Σ position = data field. Expected numbers are
computed here from the same records, independently of the engine.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts"))
from server import pivots as PV  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


RECS = [("서울", "A", "Q1", 120), ("서울", "B", "Q1", 80), ("부산", "A", "Q1", 60), ("부산", "B", "Q1", 90),
        ("대구", "A", "Q1", 40), ("대구", "B", "Q1", 30), ("서울", "A", "Q2", 130), ("서울", "B", "Q2", 70),
        ("부산", "A", "Q2", 65), ("부산", "B", "Q2", 95), ("대구", "A", "Q2", 45), ("대구", "B", "Q2", 35)]
FIELDS = ["지역", "제품", "분기", "매출"]


class Doc:
    def __init__(self, sheets):
        self.sheets = sheets


def data_sheet(recs):
    cells = {(0, c): {"v": h} for c, h in enumerate(FIELDS)}
    for r, rec in enumerate(recs, 1):
        for c, v in enumerate(rec):
            cells[(r, c)] = {"v": v}
    return {"id": "s1", "name": "데이터", "cells": cells}


def value_at(sh, r, c):
    return (sh["cells"].get((r, c)) or {}).get("v")


def expect(recs, fn, **flt):
    idx = {f: i for i, f in enumerate(FIELDS)}
    vals = [rec[3] for rec in recs if all(rec[idx[k]] == v for k, v in flt.items())]
    if not vals:
        return None
    return {"sum": sum(vals), "average": sum(vals) / len(vals), "count": len(vals)}[fn]


def grid(cells, out):
    g = {}
    for r, c, cell in out:
        g[(r, c)] = (cell or {}).get("v")
    return lambda r, c: g.get((r, c), (cells.get((r, c)) or {}).get("v"))


REG, QTR, PROD = ["대구", "부산", "서울"], ["Q1", "Q2"], ["A", "B"]

# 1) Excel compact layout: rows 지역 > 분기 with subtotals on the region row, columns 제품, grand totals
row_items = []
for i, reg in enumerate(REG):
    row_items.append({"t": "data", "r": 0, "i": 0, "x": [reg]})
    for q in QTR:
        row_items.append({"t": "data", "r": 1, "i": 0, "x": [q]})
row_items.append({"t": "grand", "r": 0, "i": 0, "x": [None]})
col_items = [{"t": "data", "r": 0, "i": 0, "x": [p]} for p in PROD] + [{"t": "grand", "r": 0, "i": 0, "x": [None]}]
lay = {"rowFields": ["지역", "분기"], "colFields": ["제품"], "rowItems": row_items, "colItems": col_items, "hidden": {},
       "firstHeaderRow": 1, "firstDataRow": 2, "firstDataCol": 1, "fields": FIELDS,
       "items": {"지역": REG, "제품": PROD, "분기": QTR}}
n_rows = len(row_items)
pivot = {"name": "P", "ref": f"A1:D{2 + n_rows}", "source": "데이터!A1:D13", "values": [{"name": "합계", "field": "매출", "fn": "sum"}], "layout": lay}
psheet = {"id": "s2", "name": "피벗", "cells": {}, "pivots": [pivot]}
doc = Doc([data_sheet(RECS), psheet])
out, notes = PV.refresh(doc, psheet, pivot, value_at)
at = grid(psheet["cells"], out)
want = {}
for k, it in enumerate(row_items):
    r = 2 + k
    reg = it["x"][0] if it["r"] == 0 and it["t"] == "data" else None
    if it["t"] == "data" and it["r"] == 0:
        cur = it["x"][0]
        flt = {"지역": cur}
    elif it["t"] == "data":
        flt = {"지역": cur, "분기": it["x"][0]}
    else:
        flt = {}
    for j, p in enumerate(PROD + [None]):
        f = dict(flt, **({"제품": p} if p else {}))
        want[(r, 1 + j)] = expect(RECS, "sum", **f)
got = {k: at(*k) for k in want}
check("Excel layout (rowItems): every data cell = the sum over its row/column items", got == want,
      str([(k, got[k], want[k]) for k in want if got[k] != want[k]][:5]))
check("…region rows are subtotals, last row/column grand totals (대구 A = 85, grand = 860)",
      at(2, 1) == 85 and at(2 + n_rows - 1, 3) == 860 and at(3, 1) == 40, f"{at(2, 1)} {at(2 + n_rows - 1, 3)} {at(3, 1)}")
check("…no notes when nothing new in the source", notes == [], str(notes))

# 2) after a source change only the affected cells change
recs2 = [list(r) for r in RECS]
recs2[0][3] = 1000  # 서울 A Q1
doc2 = Doc([data_sheet(recs2), psheet])
psheet["cells"] = {k: {"v": v, "s": {"b": True}} for k, v in want.items() if v is not None}  # as refreshed + a style
out2, _ = PV.refresh(doc2, psheet, pivot, value_at)
changed = sorted((r, c) for r, c, _ in out2)
seoul = 2 + 2 * 3  # 서울 subtotal row
check("source change: only 서울 subtotal, 서울 Q1, its totals and the grand totals change",
      changed == sorted([(seoul, 1), (seoul, 3), (seoul + 1, 1), (seoul + 1, 3), (2 + n_rows - 1, 1), (2 + n_rows - 1, 3)]), str(changed))
check("…the cell keeps its style", all((cell or {}).get("s") == {"b": True} for _, _, cell in out2))

# 3) two data fields on columns (Σ): sum and average per product
col_items2 = []
for p in PROD:
    for d in (0, 1):
        col_items2.append({"t": "data", "r": 0 if d == 0 else 1, "i": d, "x": [p] if d == 0 else [None], "d": d} if d == 0 else
                          {"t": "data", "r": 1, "i": d, "x": [None], "d": d})
lay3 = dict(lay, colFields=["제품", "Σ"], colItems=col_items2, rowFields=["지역"], rowItems=[{"t": "data", "r": 0, "i": 0, "x": [r]} for r in REG])
pivot3 = dict(pivot, ref="A1:E5", values=[{"name": "합계", "field": "매출", "fn": "sum"}, {"name": "평균", "field": "매출", "fn": "average"}], layout=lay3)
ps3 = {"id": "s3", "name": "피벗3", "cells": {}, "pivots": [pivot3]}
out3, notes3 = PV.refresh(Doc([data_sheet(RECS), ps3]), ps3, pivot3, value_at)
at3 = grid(ps3["cells"], out3)
check("two data fields on columns: sum / average per product (대구: A 85 / 42.5, B 65 / 32.5)",
      [at3(2, c) for c in range(1, 5)] == [85, 42.5, 65, 32.5], str([at3(2, c) for c in range(1, 5)]))

# 4) hidden items (a filter) leave totals without them; count
lay4 = dict(lay, rowFields=["지역"], rowItems=[{"t": "data", "r": 0, "i": 0, "x": [r]} for r in REG] + [{"t": "grand", "r": 0, "i": 0, "x": [None]}],
            hidden={"분기": ["Q2"]})
pivot4 = dict(pivot, ref="A1:D6", values=[{"name": "개수", "field": "매출", "fn": "count"}], layout=lay4)
ps4 = {"id": "s4", "name": "피벗4", "cells": {}, "pivots": [pivot4]}
out4, _ = PV.refresh(Doc([data_sheet(RECS), ps4]), ps4, pivot4, value_at)
at4 = grid(ps4["cells"], out4)
check("hidden item (분기 Q2 filtered out) + count: 1 per region/product, grand 6",
      [at4(2, 1), at4(5, 3)] == [1, 6], str([at4(2, 1), at4(5, 3)]))

# 5) a new item in the source: values for the others, a note, the grand total includes it
recs5 = RECS + [("광주", "A", "Q1", 50)]
out5, notes5 = PV.refresh(Doc([dict(data_sheet(recs5))]), psheet, dict(pivot, source="데이터!A1:D14"), value_at)
check("new item 광주: noted (needs Excel refresh), grand total includes it",
      any("광주" in n for n in notes5) and grid(psheet["cells"], out5)(2 + n_rows - 1, 3) == 910, str(notes5))

# 6) no rowItems (LibreOffice): labels in the output cells for one row field and one column field
lo_cells = {(0, 0): {"v": "Sum - 매출"}, (1, 0): {"v": "지역"}, (1, 1): {"v": "A"}, (1, 2): {"v": "B"}, (1, 3): {"v": "Total Result"}}
for k, reg in enumerate(REG):
    lo_cells[(2 + k, 0)] = {"v": reg}
lo_cells[(5, 0)] = {"v": "Total Result"}
lay6 = dict(lay, rowFields=["지역"], rowItems=None, colItems=None)
pivot6 = dict(pivot, ref="A1:D6", layout=lay6)
ps6 = {"id": "s6", "name": "피벗6", "cells": lo_cells, "pivots": [pivot6]}
out6, notes6 = PV.refresh(Doc([data_sheet(RECS), ps6]), ps6, pivot6, value_at)
at6 = grid(lo_cells, out6)
check("no rowItems: labels map rows/columns, last row/column = grand totals",
      [at6(2, 1), at6(2, 2), at6(2, 3), at6(4, 1), at6(5, 3)] == [85, 65, 150, 250, 860], str([at6(2, 1), at6(2, 2), at6(2, 3), at6(4, 1), at6(5, 3)]))

# 7) no rowItems and two row fields: not refreshed here, explained
pivot7 = dict(pivot, layout=dict(lay, rowItems=None, colItems=None))
out7, notes7 = PV.refresh(Doc([data_sheet(RECS), psheet]), psheet, pivot7, value_at)
check("no rowItems with two row fields: nothing written, note says Excel refreshes on open",
      out7 == [] and any("Excel refreshes it when it opens the file" in n for n in notes7), str(notes7))

check("aggregates: stdDev / var / product / countNums like Excel",
      round(PV.aggregate("stddev", [1, 2, 3, 4]), 6) == 1.290994 and PV.aggregate("varp", [1, 2, 3, 4]) == 1.25
      and PV.aggregate("product", [2, 3]) == 6 and PV.aggregate("countNums", [1, "a", None, 2]) == 2 and PV.aggregate("sum", ["x"]) is None)

# 8) from a file: the LibreOffice sample with Excel's rowItems/colItems added -> read -> refresh gives
#    the file's own numbers (nothing changes)
import io  # noqa: E402
import openpyxl  # noqa: E402
from openpyxl.pivot.table import RowColItem  # noqa: E402
from openpyxl.pivot.fields import Index  # noqa: E402
from server.xlsx_io import read_xlsx  # noqa: E402
wb = openpyxl.load_workbook(os.path.join(ROOT, "samples", "pivot.xlsx"))
pt = wb["피벗"]._pivots[0]
pt.rowItems = [RowColItem(x=[Index(v=k)]) for k in range(3)] + [RowColItem(t="grand", x=[Index()])]
pt.colItems = [RowColItem(x=[Index(v=k)]) for k in range(2)] + [RowColItem(t="grand", x=[Index()])]
buf = io.BytesIO(); wb.save(buf)
m = read_xlsx(buf.getvalue())
sheets = []
for sm in m["sheets"]:
    sheets.append({"id": sm["name"], "name": sm["name"], "pivots": sm.get("pivots") or [], "cells": {(r, c): it for r, c, it in sm["cells"]}})
fp = sheets[1]["pivots"][0]
check("file with rowItems: read as Excel layout (items resolved to values)",
      [x["x"] for x in fp["layout"]["rowItems"]] == [["대구"], ["부산"], ["서울"], [None]] and fp["layout"]["rowItems"][-1]["t"] == "grand",
      str(fp["layout"]["rowItems"]))
out8, notes8 = PV.refresh(Doc(sheets), sheets[1], fp, value_at)
check("…refresh reproduces the file's numbers exactly (0 changes)", out8 == [] and notes8 == [], f"{out8[:3]} {notes8}")

print(f"{fails} FAILED" if fails else "ALL PASSED")
sys.exit(1 if fails else 0)
