"""P7: every UI feature is reachable from the agent CLI.  python3 tests/agent_full.test.py"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts", "xl.py")
sys.path.insert(0, ROOT)
from openpyxl import load_workbook  # noqa: E402
from tests.make_samples import make_xlsx  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def xl(*args, ok=True):
    p = subprocess.run([sys.executable, XL, *args, "--root", ws], capture_output=True, text=True, encoding="utf-8")
    if ok and p.returncode != 0:
        print("   ", p.stdout.strip(), p.stderr.strip())
    return p


def js(*args):
    return json.loads(xl(*args, "--json").stdout)


def rows(ref, what="values"):
    return js("read", ref, "--what", what)["rows"]


ws = tempfile.mkdtemp(prefix="wx-full-")
try:
    make_xlsx(os.path.join(ws, "s.xlsx"))
    xl("serve", "--port", "8831")
    xl("open", "s.xlsx")

    # filter criteria (custom / values / clear / reapply)
    r = js("filter", "col", "B", "--where", ">150")
    check("filter custom condition hides rows", "2 rows hidden" in " ".join(r["summary"]), str(r["summary"]))
    r = js("filter", "col", "A", "--values", "3월|4월")
    check("filter by displayed values (AND with B)", "4 rows hidden" in " ".join(r["summary"]), str(r["summary"]))
    xl("filter", "clear")
    check("filter clear", json.loads(xl("describe", "--json").stdout)["sheets"][0]["autoFilter"] == "A1:F7")

    # text exactly as shown on screen
    check("read --what text applies number formats", rows("B9", "text") == [["158.50"]] and rows("F2", "text") == [["2026-01-01"]])

    # chart edit
    cid = js("describe")["sheets"][0]["charts"][0]["id"]
    xl("chart", "set", cid, "--title", "바뀐 제목", "--to-type", "line", "--legend", "b", "--at", "H20")
    ch = js("describe")["sheets"][0]["charts"][0]
    check("chart set title/type/position", ch["title"] == "바뀐 제목" and ch["type"] == "line" and ch["at"] == "H20", str(ch))

    # sheets: copy, move, hide, tab color, gridlines (active sheet must stay '매출')
    xl("sheet", "copy", "매출", "매출 복사")
    xl("sheet", "move", "매출 복사", "1")
    xl("sheet", "color", "비중", "#FF0000")
    xl("sheet", "hide", "비중")
    xl("sheet", "gridlines", "매출", "off")
    d = js("describe")
    check("sheet copy+move order", [s["name"] for s in d["sheets"]] == ["매출 복사", "매출", "비중"], str([s["name"] for s in d["sheets"]]))
    check("active sheet unchanged by copy/move", next(s for s in d["sheets"] if s["active"])["name"] == "매출")
    check("hidden sheet", d["sheets"][2]["state"] == "hidden")

    # series fill (lists, linear trend, number suffix, dates)
    xl("write", "H2", "월요일", "5", "항목1", "2026-01-31")
    xl("write", "I3", "10", "--raw")
    xl("write", "I3", "10")
    xl("series", "H2", "H2:H5")
    xl("series", "I2:I3", "I2:I5")
    xl("series", "J2", "J2:J4")
    xl("series", "K2", "K2:K3")
    check("series: weekday list", rows("H2:H5") == [["월요일"], ["화요일"], ["수요일"], ["목요일"]])
    check("series: linear trend", rows("I2:I5") == [[5], [10], [15], [20]], str(rows("I2:I5")))
    check("series: number suffix", rows("J2:J4") == [["항목1"], ["항목2"], ["항목3"]])
    check("series: date +1 day", rows("K2:K3") == [["2026-01-31"], ["2026-02-01"]])

    # copy / paste special / move
    xl("copy", "E2:E3", "M2")
    check("copy translates formulas", rows("M2:M3", "formulas") == [["=SUM(J2:L2)"], ["=SUM(J3:L3)"]], str(rows("M2:M3", "formulas")))
    xl("copy", "A1:C2", "O1", "--what", "values", "--transpose")
    check("paste values transposed", rows("O1:P3") == [["월", "1월"], ["제품A", 126], ["제품B", 104]])
    xl("copy", "A1", "R1", "--what", "formats")
    check("paste formats", (rows("R1", "all")[0][0] or {}).get("s", {}).get("bg") == "#4472C4")
    xl("move", "M2:M3", "M20")
    check("move keeps formula text and clears source", rows("M20", "formulas") == [["=SUM(J2:L2)"]] and rows("M2") == [[None]])

    # insert / delete cells (shift)
    xl("write", "T1", "x")
    xl("cells", "insert", "T1", "--shift", "down")
    moved = rows("T1:T2")
    xl("cells", "delete", "T1", "--shift", "up")
    check("cells insert(down)/delete(up)", moved == [[None], ["x"]] and rows("T1") == [["x"]], f"{moved} {rows('T1')}")

    # remove duplicates
    xl("write", "V1", "이름", "--down")
    xl("write", "V2", "a", "b", "a", "c", "--down")
    r = js("dedupe", "V1:V5")
    check("remove duplicates", rows("V1:V4") == [["이름"], ["a"], ["b"], ["c"]] and "removed 1 duplicate rows" in " ".join(r["summary"]))

    # find / replace
    f = js("find", "*월", "--max", "50")
    check("find with wildcard", any(m["cell"] == "A2" for m in f["matches"]))
    xl("replace", "제품", "품목", "--all")
    check("replace across sheets", rows("B1") == [["품목A"]] and rows("비중!A1") == [["품목"]])
    xl("replace", "SUM", "SUM", "--in", "formulas")

    # comment, hyperlink, autofit, format copy, decimals, merge across
    xl("comment", "B2", "확인 필요")
    xl("link", "A13", "https://example.com", "--text", "사이트")
    xl("width", "A", "3")
    xl("autofit", "A:A")
    xl("format-copy", "A1", "X1:X3")
    xl("decimals", "B9", "1")
    xl("merge", "Z1:AA3", "--across")
    check("comment set", rows("B2", "all")[0][0].get("cm") == "확인 필요")
    check("hyperlink set", rows("A13", "all")[0][0].get("h") == "https://example.com")
    check("decimals +1", rows("B9", "text") == [["158.500"]])
    check("merge across = one merge per row", js("describe")["sheets"][1]["merges"] == 3)

    # goto / view (no browser connected -> still succeeds)
    check("goto", xl("goto", "B5").returncode == 0)
    check("view", xl("view", "--zoom", "150").returncode == 0)

    xl("save", "--as", "out.xlsx")
    wb = load_workbook(os.path.join(ws, "out.xlsx"))
    sh = wb["매출"]
    check("saved: sheet order", wb.sheetnames == ["매출 복사", "매출", "비중"])
    check("saved: hidden sheet + tab color", wb["비중"].sheet_state == "hidden" and wb["비중"].sheet_properties.tabColor.rgb.endswith("FF0000"))
    check("saved: gridlines off", sh.sheet_view.showGridLines is False)
    check("saved: comment", sh["B2"].comment is not None and sh["B2"].comment.text.startswith("확인 필요"))
    check("saved: hyperlink", sh["A13"].hyperlink is not None and sh["A13"].hyperlink.target == "https://example.com")
    check("saved: column autofit wider than 3 chars", (sh.column_dimensions["A"].width or 0) > 5)
    check("saved: merges", {"Z1:AA1", "Z2:AA2", "Z3:AA3"} <= {str(m) for m in sh.merged_cells.ranges})
    check("saved: chart title", len(sh._charts) == 1)
finally:
    xl("stop", "--force", ok=False)
    shutil.rmtree(ws, ignore_errors=True)

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
