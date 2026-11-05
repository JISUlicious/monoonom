"""Charts addressed by id work on any sheet: python3 tests/chart_ids.test.py

`chart ls` lists the charts of every sheet, but `chart set/rm ID` used to look only in the active
sheet, so an id from another sheet gave 404. Also: `--at 'Sheet!E2'` used to be a server error.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "xl.py")
fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def xl(*args):
    return subprocess.run([sys.executable, XL, *args, "--root", ws], capture_output=True, text=True, encoding="utf-8")


def charts():
    """{sheet name: [(id, title, at)]}"""
    d = json.loads(xl("describe", "--json").stdout)
    return {s["name"]: [(c["id"], c["title"], c["at"]) for c in s["charts"]] for s in d["sheets"]}


ws = tempfile.mkdtemp(prefix="wx-charts-")
try:
    shutil.copyfile(os.path.join(ROOT, "samples", "sample.xlsx"), os.path.join(ws, "s.xlsx"))
    xl("serve", "--port", "8834")
    xl("open", "s.xlsx")
    c = charts()
    check("sample: one chart on 매출 (active sheet), one on 비중", len(c["매출"]) == 1 and len(c["비중"]) == 1, str(c))
    pie = c["비중"][0][0]

    p = xl("chart", "set", pie, "--title", "새 제목")
    check("set an id from a non-active sheet without --sheet", p.returncode == 0 and charts()["비중"][0][1] == "새 제목", p.stderr.strip())
    p = xl("chart", "set", pie, "--data", "비중!A1:B3")
    check("…also with --data", p.returncode == 0, p.stderr.strip())
    v = json.loads(xl("changes", "--since", "0", "--json").stdout)["changes"][-1]
    check("…and only that sheet is touched (no no-op changes for the other sheets searched)", v["meta"] == ["비중: charts"], str(v.get("meta")))

    p = xl("chart", "set", pie, "--title", "x", "--sheet", "매출")
    check("--sheet limits the search and says where it looked", p.returncode == 5 and "'매출' 시트" in p.stderr, p.stderr.strip())
    p = xl("chart", "rm", "chNOPE")
    check("unknown id: 404 that says every sheet was searched", p.returncode == 5 and "모든 시트" in p.stderr, p.stderr.strip())

    p = xl("chart", "add", "매출!A1:B7", "--type", "line", "--at", "비중!E12")
    c = charts()
    check("add --at 'Sheet!E12' puts the chart on that sheet", p.returncode == 0 and any(at == "E12" for _, _, at in c["비중"]), p.stderr.strip() or str(c))
    p = xl("chart", "add", "매출!A1:B7", "--sheet", "매출", "--at", "비중!E30")
    check("--sheet and a different --at sheet: clear 400", p.returncode != 0 and "--sheet와 --at" in p.stderr, p.stderr.strip())
    p = xl("chart", "add", "매출!A1:B7", "--at", "A0")
    check("bad --at: 400 with the reason, not a server error", p.returncode != 0 and "차트 위치" in p.stderr and "500" not in p.stderr, p.stderr.strip())

    line = next(i for i, _, at in charts()["비중"] if at == "E12")
    p = xl("chart", "set", line, "--at", "매출!K30")
    c = charts()
    check("set --at 'Other!K30' moves the chart to that sheet", p.returncode == 0 and any(i == line and at == "K30" for i, _, at in c["매출"]) and all(i != line for i, _, _ in c["비중"]), str(c))

    p = xl("chart", "rm", pie)
    check("rm an id from a non-active sheet without --sheet", p.returncode == 0 and all(i != pie for i, _, _ in charts()["비중"]), p.stderr.strip())

    xl("save", "--as", "out.xlsx")
    from openpyxl import load_workbook
    wb = load_workbook(os.path.join(ws, "out.xlsx"))
    check("saved: moved chart on 매출, pie gone from 비중", len(wb["매출"]._charts) == 2 and len(wb["비중"]._charts) == 0,
          f"매출={len(wb['매출']._charts)} 비중={len(wb['비중']._charts)}")
finally:
    xl("stop", "--force")
    shutil.rmtree(ws, ignore_errors=True)

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
