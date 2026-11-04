"""P3 server-side structural edits and conflict transform:  python3 tests/structural.test.py"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "_vendor"))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts"))
sys.path.insert(0, ROOT)
from server.agent_api import build_ops, structural_batch  # noqa: E402
from server.session import DocSession, SessionError  # noqa: E402
from server.app import read_workbook  # noqa: E402
from tests.make_samples import make_xlsx  # noqa: E402
import tempfile  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def fresh():
    p = os.path.join(tempfile.mkdtemp(), "s.xlsx")
    make_xlsx(p)
    return DocSession(read_workbook(p), p)


def agent(doc, edit):
    ops, shifts, transformable, label = structural_batch(doc, edit)
    return doc.apply(ops, actor="agent", client_id="cli", structural=True, shifts=shifts, transformable=transformable, label=label)


def f(doc, sheet, r, c):
    return (doc.sheet(sheet)["cells"].get((r, c)) or {}).get("f")


def v(doc, sheet, r, c):
    return (doc.sheet(sheet)["cells"].get((r, c)) or {}).get("v")


# --- insert / delete rows and columns
d = fresh()
agent(d, {"op": "insertRows", "sheet": "매출", "at": "3", "count": 2})
check("insert rows moves data", v(d, "매출", 4, 0) == "2월" and v(d, "매출", 2, 0) is None)
check("insert rows adjusts totals", f(d, "매출", 9, 1) == "=SUM(B2:B9)", f(d, "매출", 9, 1))
check("insert rows adjusts other sheet", f(d, "비중", 0, 3) == "=매출!E10", f(d, "비중", 0, 3))
check("insert rows extends filter", d.sheet("매출")["autoFilter"]["ref"] == "A1:F9")
check("insert rows extends chart series", d.sheet("매출")["charts"][0]["series"][0]["val"] == "'매출'!$B$2:$B$9"
      or d.sheet("매출")["charts"][0]["series"][0]["val"] == "매출!$B$2:$B$9", d.sheet("매출")["charts"][0]["series"][0]["val"])
agent(d, {"op": "deleteRows", "sheet": "매출", "at": "3:4"})
check("delete rows restores", v(d, "매출", 2, 0) == "2월" and f(d, "매출", 7, 1) == "=SUM(B2:B7)")
agent(d, {"op": "deleteRows", "sheet": "매출", "at": "2"})
check("deleting a referenced row shrinks ranges", f(d, "매출", 6, 1) == "=SUM(B2:B6)", f(d, "매출", 6, 1))
agent(d, {"op": "insertCols", "sheet": "비중", "at": "A"})
check("insert column on other sheet keeps its refs", f(d, "비중", 0, 4) == "=매출!E7", f(d, "비중", 0, 4))
check("merge shifts with columns", d.sheet("비중")["merges"] == ["B6:C6"], str(d.sheet("비중")["merges"]))

# --- sort
d = fresh()
agent(d, {"op": "sort", "range": "매출!A1:F7", "by": ["B:desc"]})
check("sort desc keeps header", v(d, "매출", 0, 0) == "월" and v(d, "매출", 1, 0) == "6월", v(d, "매출", 1, 0))
check("sorted formulas follow their row", f(d, "매출", 1, 4) == "=SUM(B2:D2)")
check("totals untouched", f(d, "매출", 7, 1) == "=SUM(B2:B7)")

# --- conflict transform: a client edit based on an old version, after an agent row insert
d = fresh()
sid = d.sheet("매출")["id"]
base = d.version
agent(d, {"op": "insertRows", "sheet": "매출", "at": "2", "count": 3})
d.apply([{"op": "setCells", "sheet": sid, "cells": [[5, 6, {"v": "사용자"}], [6, 6, {"f": "=B7*2"}]]}],
        client_id="browser", base_version=base)
check("stale cell edit moved down", v(d, "매출", 8, 6) == "사용자" and v(d, "매출", 5, 6) is None)
check("stale formula adjusted", f(d, "매출", 9, 6) == "=B10*2", f(d, "매출", 9, 6))
check("entry marked transformed", d.log[-1]["transformed"] is True)

base = d.version
agent(d, {"op": "deleteRows", "sheet": "매출", "at": "2:4"})
d.apply([{"op": "setCells", "sheet": sid, "cells": [[2, 7, {"v": "삭제될 행"}], [9, 7, {"v": "남는 행"}]]}],
        client_id="browser", base_version=base)
check("edit in deleted row dropped, later row moved up", v(d, "매출", 6, 7) == "남는 행" and v(d, "매출", 2, 7) != "삭제될 행")

base = d.version
agent(d, {"op": "sort", "range": "매출!A1:F7", "by": ["C"]})
try:
    d.apply([{"op": "setCells", "sheet": sid, "cells": [[3, 3, {"v": 1}]]}], client_id="browser", base_version=base)
    check("stale edit after sort rejected", False)
except SessionError as e:
    check("stale edit after sort rejected", e.status == 409)

base = d.version
agent(d, {"op": "insertCols", "sheet": "매출", "at": "A"})
try:
    d.apply([{"op": "setSheetMeta", "sheet": sid, "meta": {"cols": {}}}], client_id="browser", base_version=base)
    check("stale sheet-properties change on shifted sheet rejected", False)
except SessionError as e:
    check("stale sheet-properties change on shifted sheet rejected", e.status == 409)

# --- undo a structural agent batch
d = fresh()
agent(d, {"op": "insertRows", "sheet": "매출", "at": "3", "count": 2})
batch = d.log[-1]
d.apply(batch["undo"], actor="agent", structural=True, shifts=[{**x, "count": -x["count"]} for x in batch["shifts"]], transformable=True)
check("undo insert restores cells and formulas", v(d, "매출", 2, 0) == "2월" and f(d, "매출", 7, 1) == "=SUM(B2:B7)" and f(d, "비중", 0, 3) == "=매출!E8")

# --- mixed batch through the edit splitter keeps coordinates consistent
d = fresh()
ops, _, _ = build_ops(d, [{"op": "set", "range": "매출!G1", "values": [["메모"]]}])
d.apply(ops, actor="agent")
agent(d, {"op": "insertRows", "sheet": "매출", "at": "1"})
check("header row insert moves earlier write", v(d, "매출", 1, 6) == "메모")

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
