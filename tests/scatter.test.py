"""Scatter charts: python3 tests/scatter.test.py   (starts its own server, port 8842)

- the agent's `chart add --type scatter` takes a numeric first column as X (as the UI and Excel)
- saved scatter charts have the required <scatterStyle> (ECMA-376 CT_ScatterChart)
- axis positions are written: X / category axis at the bottom, values left (horizontal bars: swapped)
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "xl.py")
fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def xl(*args):
    p = subprocess.run([sys.executable, XL, *args, "--root", ws], capture_output=True, text=True, encoding="utf-8")
    if p.returncode != 0:
        print("   ", p.stdout.strip(), p.stderr.strip())
    return p


def charts_xml(path):
    z = zipfile.ZipFile(path)
    out = {}
    for n in sorted(x for x in z.namelist() if x.startswith("xl/charts/chart")):
        s = z.read(n).decode("utf-8")
        t = re.findall(r"<a:t>([^<]*)", s)
        out[t[0] if t else n] = s
    return out


def ax_pos(s, tag):
    """axPos of every <tag> axis in chart XML s."""
    return [re.search(r'<(?:\w+:)?axPos val="(\w)"', m).group(1)
            for m in re.findall(rf"<(?:\w+:)?{tag}>.*?</(?:\w+:)?{tag}>", s, re.S)]


ws = tempfile.mkdtemp(prefix="wx-scatter-")
try:
    xl("serve", "--port", "8842")
    xl("new", "--name", "t")
    xl("write", "A1", "--data", json.dumps([["키", "몸무게", "나이"], [160, 55, 20], [170, 65, 30], [180, 80, 40], [175, 70, 35]]), "--doc", "t")
    xl("write", "E1", "--data", json.dumps([["이름", "점수"], ["가", 3], ["나", 5], ["다", 4]]), "--doc", "t")
    xl("chart", "add", "A1:C5", "--type", "scatter", "--title", "num", "--doc", "t")
    xl("chart", "add", "E1:F4", "--type", "scatter", "--title", "txt", "--at", "H20", "--doc", "t")
    xl("chart", "add", "A1:C5", "--title", "col", "--at", "P2", "--doc", "t")
    xl("chart", "add", "A1:C5", "--type", "barH", "--title", "barh", "--at", "P20", "--doc", "t")
    d = json.loads(xl("describe", "--json", "--doc", "t").stdout)
    ch = {c["title"]: c for c in d["sheets"][0]["charts"]}
    s = ch["num"]["series"]
    check("agent scatter: the numeric first column is X, not a series (2 series: 몸무게, 나이)",
          len(s) == 2 and all(x["x"].endswith("$A$2:$A$5") for x in s) and s[0]["y"].endswith("$B$2:$B$5"), json.dumps(s, ensure_ascii=False))
    t = ch["txt"]["series"]
    check("agent scatter with a text first column: that column is X (categories)", len(t) == 1 and t[0]["x"].endswith("$E$2:$E$4"), json.dumps(t, ensure_ascii=False))

    xl("save", "--as", "out.xlsx", "--doc", "t")
    x = charts_xml(os.path.join(ws, "out.xlsx"))
    check("saved scatter: <scatterStyle val=\"lineMarker\"> (required element)",
          all(re.search(r'<(?:\w+:)?scatterStyle val="lineMarker"', x[k]) for k in ("num", "txt")))
    check("…markers only: series lines are noFill", all("<a:noFill" in x[k] and "<marker>" in x[k].replace("c:", "") for k in ("num", "txt")))
    check("scatter axes: X bottom, Y left", ax_pos(x["num"], "valAx") == ["b", "l"], str(ax_pos(x["num"], "valAx")))
    check("column chart: categories bottom, values left", ax_pos(x["col"], "catAx") == ["b"] and ax_pos(x["col"], "valAx") == ["l"],
          str((ax_pos(x["col"], "catAx"), ax_pos(x["col"], "valAx"))))
    check("horizontal bar: categories left, values bottom", ax_pos(x["barh"], "catAx") == ["l"] and ax_pos(x["barh"], "valAx") == ["b"],
          str((ax_pos(x["barh"], "catAx"), ax_pos(x["barh"], "valAx"))))
    # LibreOffice reads the file back as scatter with 2 series (when installed)
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if lo:
        subprocess.run([lo, "--headless", "--convert-to", "pdf", "--outdir", ws, os.path.join(ws, "out.xlsx")], capture_output=True, timeout=120)
        check("LibreOffice opens the saved file", os.path.exists(os.path.join(ws, "out.pdf")))
finally:
    xl("stop", "--force")
    shutil.rmtree(ws, ignore_errors=True)
print(f"{fails} FAILED" if fails else "ALL PASSED")
sys.exit(1 if fails else 0)
