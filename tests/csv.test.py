"""CSV / TSV: open as one sheet (encoding and delimiter detected, values read like typing), a plain
save writes NAME.xlsx and leaves the .csv alone, `save --as x.csv` writes the active sheet's values.
    python3 tests/csv.test.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts", "xl.py")
fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra and not cond else ""))
    if not cond:
        fails += 1


ws = tempfile.mkdtemp(prefix="wx-csv-")


def xl(*a, ok=True):
    p = subprocess.run([sys.executable, XL, *a, "--root", ws], capture_output=True, text=True)
    if ok and p.returncode:
        raise SystemExit(f"xl {' '.join(a)} failed: {p.stderr}")
    return p


open(os.path.join(ws, "sales.csv"), "w", encoding="utf-8-sig", newline="").write(
    'Item,Qty,Price,Rate,Date,Code,Total\r\nApple,3,"1,200",12%,2026-03-15,00123,=B2*C2\r\nPear,5,800,5%,2026-03-16,A7,=B3*C3\r\n')
open(os.path.join(ws, "kr.csv"), "wb").write("이름;수량\n사과;3\n배;5\n".encode("cp949"))
open(os.path.join(ws, "t.tsv"), "w", encoding="utf-8").write("a\tb\n1\t2\n")
try:
    xl("serve", "--port", "8882")
    check("ls lists CSV and TSV files", all(n in xl("ls").stdout for n in ("sales.csv", "kr.csv", "t.tsv")))
    out = xl("open", "sales.csv").stdout
    check("open: one sheet named after the file, with a note about saving", "sheets: sales" in out and "opened as one sheet" in out and "utf-8" in out, out)
    r = json.loads(xl("read", "A1:G3", "--what", "all", "--json").stdout)
    rows = r["rows"]
    vals = [[c.get("v") if isinstance(c, dict) else c for c in row] for row in rows]
    check("values read like typing: numbers, 1,200 -> 1200, 12% -> 0.12, a date, 00123 -> 123 (as Excel)",
          vals[1][1] == 3 and vals[1][2] == 1200 and abs(vals[1][3] - 0.12) < 1e-9 and vals[1][5] == 123 and vals[2][5] == "A7", json.dumps(rows[1:])[:400])
    v = json.loads(xl("read", "E2:G3", "--json").stdout)["rows"]
    check("…the date shows as a date, = makes a formula (computed)", v[0][0] == "2026-03-15" and v[0][2] == 3600 and v[1][2] == 4000, json.dumps(v))
    xl("write", "H1", "Note")
    p = xl("save")
    check("plain save writes sales.xlsx and says the .csv is unchanged", os.path.isfile(os.path.join(ws, "sales.xlsx")) and "The original .csv is unchanged" in p.stdout, p.stdout)
    check("…the .csv itself is not touched", "Note" not in open(os.path.join(ws, "sales.csv"), encoding="utf-8-sig").read())
    from openpyxl import load_workbook
    wb = load_workbook(os.path.join(ws, "sales.xlsx"))
    check("…the .xlsx keeps the formula and the new cell", wb.active["G2"].value == "=B2*C2" and wb.active["H1"].value == "Note")
    p = xl("save", "--as", "out.csv")
    data = open(os.path.join(ws, "out.csv"), "rb").read()
    text = data.decode("utf-8-sig")
    check("save --as out.csv: UTF-8 with BOM, formula results and dates as text, warns what is lost",
          data.startswith(b"\xef\xbb\xbf") and "Apple,3,1200,0.12,2026-03-15,123,3600,\r\n".replace(",\r\n", "\r\n") in text
          and "formulas (their results are saved)" in p.stdout, repr(text[:200]) + p.stdout)
    out = xl("open", "kr.csv").stdout
    k = json.loads(xl("read", "A1:B3", "--json", "--doc", "kr.csv").stdout)["rows"]
    check("CP949 file with ; is read (Korean Excel's CSV)", "cp949" in out and k == [["이름", "수량"], ["사과", 3], ["배", 5]], json.dumps(k, ensure_ascii=False))
    xl("open", "t.tsv")
    t = json.loads(xl("read", "A1:B2", "--json", "--doc", "t.tsv").stdout)["rows"]
    check("TSV is read with tabs", t == [["a", "b"], [1, 2]], json.dumps(t))
    p = xl("open", "nope.txt", ok=False)
    check("other extensions: refused, naming csv / tsv as supported", p.returncode != 0 and "csv / tsv" in p.stderr, p.stderr)
finally:
    xl("stop", "--force", ok=False)
    shutil.rmtree(ws, ignore_errors=True)

print("ALL PASSED" if not fails else f"{fails} FAILED")
sys.exit(1 if fails else 0)
