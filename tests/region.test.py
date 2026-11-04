"""Tables that do not start at A1 (title, blank rows, notes, totals row):  python3 tests/region.test.py"""
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "_vendor"))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts"))
sys.path.insert(0, ROOT)
from server import region as RG  # noqa: E402
from server.agent_api import build_ops, structural_batch  # noqa: E402
from server.app import read_workbook  # noqa: E402
from server.session import DocSession  # noqa: E402
from tests.make_samples import make_offset_xlsx  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def fresh():
    p = os.path.join(tempfile.mkdtemp(), "o.xlsx")
    make_offset_xlsx(p)
    return DocSession(read_workbook(p), p)


def edit(d, *edits):
    ops, structural, _ = build_ops(d, list(edits))
    d.apply(ops, actor="agent", structural=structural, transformable=build_ops.last_transformable)


A1 = lambda g: f"{chr(65 + g[1])}{g[0] + 1}:{chr(65 + g[3])}{g[2] + 1}"

d = fresh()
sh = d.sheet("지점")
check("current region skips title and notes column gap", A1(RG.current_region(sh, 6, 3)) == "C5:F13", A1(RG.current_region(sh, 6, 3)))
check("totals row detected and trimmed", A1(RG.table_body(sh, RG.current_region(sh, 6, 3))) == "C5:F12")
tabs = [A1(t) for t in RG.tables(sh)]
check("tables: data table first, title and note separate", tabs[0] == "C5:F13" and "B2:B2" in tabs and "H5:H5" in tabs, str(tabs))

edit(d, {"op": "filter", "range": "D7"})
check("filter from one cell -> table without totals", d.sheet("지점")["autoFilter"]["ref"] == "C5:F12")
edit(d, {"op": "filter", "range": None})
edit(d, {"op": "filter", "range": "C:F"})
check("filter from whole columns -> table", d.sheet("지점")["autoFilter"]["ref"] == "C5:F12")
edit(d, {"op": "filterColumn", "column": "D", "custom": [{"op": ">", "val": "100"}]})
check("criteria relative to the table's first column", d.sheet("지점")["filterHidden"] == [6, 9, 11], str(d.sheet("지점")["filterHidden"]))
edit(d, {"op": "filterClear"})

ops, shifts, tr, label = structural_batch(d, {"op": "sort", "range": "D7", "by": ["D:desc"]})
d.apply(ops, actor="agent", structural=True, transformable=tr)
cells = d.sheet("지점")["cells"]
check("sort from one cell keeps header", cells[(4, 2)]["v"] == "지점")
check("sort from one cell sorts data", cells[(5, 2)]["v"] == "대구", cells[(5, 2)]["v"])
check("totals row untouched by sort", cells[(12, 2)]["v"] == "합계" and cells[(12, 3)]["f"] == "=SUM(D6:D12)")

d = fresh()
edit(d, {"op": "filterReapply"})  # no filter yet: no-op
edit(d, {"op": "filter", "range": "C5:F12"}, {"op": "set", "range": "C13", "values": [["합계"]]})
edit(d, {"op": "filterReapply"})
check("reapply does not swallow the totals row", d.sheet("지점")["autoFilter"]["ref"] == "C5:F12", d.sheet("지점")["autoFilter"]["ref"])

d = fresh()
edit(d, {"op": "removeDuplicates", "range": "C7"})
check("remove duplicates from one cell uses the table", [d.sheet("지점")["cells"].get((r, 2), {}).get("v") for r in range(5, 12)] ==
      ["서울", "부산", "대구", "광주", "대전", "울산", None])
check("totals row kept by remove duplicates", d.sheet("지점")["cells"][(12, 2)]["v"] == "합계")

d = fresh()
edit(d, {"op": "chart", "data": "D7", "type": "bar"})
ch = d.sheet("지점")["charts"][0]
check("chart from one cell: categories = 지점 column, no totals", ch["series"][0]["cat"].endswith("$C$6:$C$12") and len(ch["series"]) == 3,
      str(ch["series"][0]))

explicit = fresh()
edit(explicit, {"op": "filter", "range": "C5:F13"})
check("explicit range is respected (totals included on purpose)", explicit.sheet("지점")["autoFilter"]["ref"] == "C5:F13")

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
