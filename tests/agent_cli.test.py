"""P2 agent CLI scenario (no browser): python3 tests/agent_cli.test.py

Runs xl.py in a temporary workspace exactly as an agent would and checks the saved file with openpyxl.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "xl.py")
sys.path.insert(0, ROOT)
from openpyxl import load_workbook  # noqa: E402
from tests.make_samples import make_xlsx  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def xl(ws, *args, ok=True):
    p = subprocess.run([sys.executable, XL, *args, "--root", ws], capture_output=True, text=True, encoding="utf-8")
    if ok and p.returncode != 0:
        print(p.stdout, p.stderr)
    return p


def js(ws, *args):
    return json.loads(xl(ws, *args, "--json").stdout)


ws = tempfile.mkdtemp(prefix="wx-cli-")
try:
    make_xlsx(os.path.join(ws, "sales.xlsx"))
    p = xl(ws, "serve", "--port", "8811")
    check("serve starts", p.returncode == 0 and "새로 시작" in p.stdout, p.stdout.strip())
    check("serve again reuses", "재사용" in xl(ws, "serve").stdout)

    check("open", "시트: 매출, 비중" in xl(ws, "open", "sales.xlsx").stdout)
    d = js(ws, "describe")
    check("describe header", d["sheets"][0]["header"][:3] == ["월", "제품A", "제품B"], str(d["sheets"][0].get("header")))
    check("describe column types", d["sheets"][0]["columnTypes"]["F"] == "date")

    r = js(ws, "read", "매출!A1:B3")
    check("read values", r["rows"] == [["월", "제품A"], ["1월", 126], ["2월", 139]], str(r["rows"]))
    check("read dates as ISO", js(ws, "read", "F2")["rows"] == [["2026-01-01"]])

    xl(ws, "write", "G1", "비율")
    xl(ws, "fill", "G2:G7", "=B2/E2")
    f = js(ws, "read", "G2:G3", "--what", "formulas")["rows"]
    check("fill shifts relative refs", f == [["=B2/E2"], ["=B3/E3"]], str(f))
    v = js(ws, "read", "G2")
    check("formula computed without browser", isinstance(v["rows"][0][0], float) and v["calc"]["status"] == "fresh",
          json.dumps(v["calc"], ensure_ascii=False))

    check("parse like Excel input", js(ws, "write", "H2", "2026-03-15", "15%", "1,234", "TRUE")["version"] > 0)
    check("parsed values", js(ws, "read", "H2:K2", "--what", "raw")["rows"] == [[46096, 0.15, 1234, True]])

    res = js(ws, "format", "G1:G7", "nf=0.0%", "b=true", "--border", "outline")
    batch = res["batchId"]
    st = js(ws, "read", "G2", "--what", "all")["rows"][0][0]["s"]
    check("format applied", st.get("nf") == "0.0%" and st.get("b") is True and "bl" in st, str(st))
    xl(ws, "undo", batch)
    st = (js(ws, "read", "G2", "--what", "all")["rows"][0][0] or {}).get("s") or {}
    check("undo agent batch", "nf" not in st, str(st))

    xl(ws, "width", "A", "16")
    xl(ws, "chart", "add", "A1:D7", "--type", "line", "--title", "추이", "--at", "H10")
    check("chart listed", "line" in xl(ws, "chart", "ls").stdout)
    xl(ws, "sheet", "rename", "매출", "판매")
    check("rename rewrites other sheets' formulas", js(ws, "read", "비중!D1", "--what", "formulas")["rows"] == [["=판매!E8"]])
    xl(ws, "sheet", "add", "요약")
    xl(ws, "write", "요약!A1", "=SUM(판매!B2:B7)")
    check("cross-sheet formula", js(ws, "read", "요약!A1")["rows"] == [[951]])

    # documented batch example (references/edits.md) on a fresh sheet: later edits see earlier cells
    xl(ws, "sheet", "add", "예제")
    ex = open(os.path.join(ROOT, ".agents", "skills", "web-excel", "references", "edits.md"), encoding="utf-8").read()
    edits = json.loads(ex.split("```json")[1].split("```")[0].replace("Sheet1!", "예제!"))
    for e in edits:
        e.setdefault("sheet", "예제")
    with open(os.path.join(ws, "e.json"), "w", encoding="utf-8") as fh:
        json.dump(edits, fh, ensure_ascii=False)
    p = xl(ws, "apply", "e.json")
    check("apply: documented example", p.returncode == 0, (p.stdout + p.stderr).strip()[:200])
    hdr = js(ws, "read", "예제!A1:C1", "--what", "all")["rows"][0]
    check("apply: format covers A1:C1", all((c or {}).get("s", {}).get("b") for c in hdr), str(hdr))
    check("apply: chart uses batch data", "$B$2:$B$3" in xl(ws, "chart", "ls").stdout)

    # P3: structural edits
    xl(ws, "rows", "insert", "3", "--count", "2", "--sheet", "판매")
    check("rows insert shifts totals", js(ws, "read", "판매!B10", "--what", "formulas")["rows"] == [["=SUM(B2:B9)"]])
    check("rows insert updates other sheet", js(ws, "read", "비중!D1", "--what", "formulas")["rows"] == [["=판매!E10"]])
    xl(ws, "rows", "delete", "3:4", "--sheet", "판매")
    xl(ws, "sort", "판매!A1:F7", "--by", "B:desc")
    check("sort desc", js(ws, "read", "판매!A2:B2")["rows"] == [["6월", 191]], str(js(ws, "read", "판매!A2:B2")["rows"]))
    xl(ws, "cols", "insert", "B", "--sheet", "판매")
    check("cols insert", js(ws, "read", "판매!C1")["rows"] == [["제품A"]])
    xl(ws, "cols", "delete", "B", "--sheet", "판매")

    ch = js(ws, "changes", "--since", "0")
    check("change log has agent entries", len(ch["changes"]) >= 8 and all(c["actor"] == "agent" for c in ch["changes"]))

    p = xl(ws, "save")
    check("save with backup", p.returncode == 0 and "백업" in p.stdout, p.stdout.strip())
    wb = load_workbook(os.path.join(ws, "sales.xlsx"))
    check("saved: sheet names", wb.sheetnames == ["판매", "비중", "요약", "예제"], str(wb.sheetnames))
    check("saved: formulas", wb["판매"]["G3"].value == "=B3/E3" and wb["비중"]["D1"].value == "=판매!E8", f'{wb["판매"]["G3"].value} {wb["비중"]["D1"].value}')
    check("saved: charts", len(wb["판매"]._charts) == 2)
    check("saved: column width", round(wb["판매"].column_dimensions["A"].width, 1) == 16.0)
    check("saved: date", wb["판매"]["H2"].is_date)

    # .xls is read-only: saved as NAME.xlsx; an existing NAME.xlsx is named before it is replaced
    from tests.make_samples import make_xls
    make_xls(os.path.join(ws, "legacy.xls"))
    shutil.copyfile(os.path.join(ws, "sales.xlsx"), os.path.join(ws, "legacy.xlsx"))
    xl(ws, "open", "legacy.xls")
    xl(ws, "write", "A6", "추가", "--doc", "legacy.xls")
    p = xl(ws, "save", "--doc", "legacy.xls", ok=False)
    check("xls save: existing xlsx is named in the 409", p.returncode == 4 and "'legacy.xlsx'" in p.stderr and "덮어쓰게" in p.stderr, p.stderr.strip())
    p = xl(ws, "save", "--doc", "legacy.xls", "--overwrite")
    check("xls save --overwrite: warns that .xls formulas are values", p.returncode == 0 and "계산된 값으로만" in p.stdout, p.stdout.strip())
    check("xls saved next to it, original untouched",
          load_workbook(os.path.join(ws, "legacy.xlsx"))["데이터"]["A6"].value == "추가" and os.path.isfile(os.path.join(ws, "legacy.xls")))
    xl(ws, "close", "--doc", "legacy.xlsx", "--force")

    check("bad sheet -> exit 5", xl(ws, "read", "없음!A1", ok=False).returncode == 5)
    check("bad doc -> exit 2", xl(ws, "read", "A1", "--doc", "zzz", ok=False).returncode == 2)
    xl(ws, "write", "A30", "미저장")
    check("stop refuses with unsaved changes", xl(ws, "stop", ok=False).returncode == 4)
    check("stop --force", xl(ws, "stop", "--force").returncode == 0)
    check("status after stop", xl(ws, "status", ok=False).returncode == 3)
finally:
    xl(ws, "stop", "--force", ok=False)
    shutil.rmtree(ws, ignore_errors=True)

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
