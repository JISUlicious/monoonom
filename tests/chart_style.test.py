"""Charts look the same in the preview and in Excel: python3 tests/chart_style.test.py

Charts web-excel creates or edits are saved with explicit formatting from server/chart_style.py
(Excel 2016+ default look), colors come from the workbook's theme or the chart's own colors, and new
workbooks get the Office 2013-2022 theme instead of openpyxl's Office 2007 one.
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
SK = os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts")
XL = os.path.join(SK, "xl.py")
sys.path.insert(0, SK)
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, PieChart, Reference  # noqa: E402
from openpyxl.chart.marker import DataPoint  # noqa: E402
from openpyxl.chart.shapes import GraphicalProperties  # noqa: E402
from openpyxl.drawing.colors import ColorChoice, SchemeColor  # noqa: E402

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


def attrs(s, tag, attr="val"):
    return re.findall(rf"<(?:\w+:)?{tag}\b[^>]*\b{attr}=\"([^\"]*)\"", s)


ws = tempfile.mkdtemp(prefix="wx-cstyle-")
try:
    # a workbook (Office 2007 theme, as openpyxl writes) whose charts have explicit colors
    wb = Workbook()
    sh = wb.active
    sh.title = "S"
    for r in [["m", "A", "B"], ["1", 3, 4], ["2", 5, 2], ["3", 4, 6]]:
        sh.append(r)
    b = BarChart()
    b.title = "colored"
    b.add_data(Reference(sh, min_col=2, max_col=3, min_row=1, max_row=4), titles_from_data=True)
    b.series[0].graphicalProperties = GraphicalProperties(solidFill="C00000")
    b.series[1].graphicalProperties = GraphicalProperties(solidFill=ColorChoice(schemeClr=SchemeColor(val="accent6", lumMod=75000)))
    sh.add_chart(b, "E2")
    p = PieChart()
    p.title = "slices"
    p.add_data(Reference(sh, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    p.series[0].dPt = [DataPoint(idx=0, spPr=GraphicalProperties(solidFill="112233")),
                       DataPoint(idx=1, spPr=GraphicalProperties(solidFill=ColorChoice(schemeClr=SchemeColor(val="accent2"))))]
    sh.add_chart(p, "E20")
    wb.save(os.path.join(ws, "old.xlsx"))

    xl("serve", "--port", "8836")

    # ---- 1) new workbook: modern theme, new charts carry the whole style
    xl("new", "--name", "fresh")
    # more rows than series columns, so the series are the columns (Excel's rule)
    xl("write", "A1", "--data", '[["월","A","B","C"],["1월",3,4,5],["2월",5,2,4],["3월",4,6,3],["4월",6,5,2]]', "--doc", "fresh")
    xl("chart", "add", "A1:D5", "--title", "막대", "--legend", "b", "--doc", "fresh")
    xl("chart", "add", "A1:B3", "--type", "line", "--doc", "fresh")
    xl("chart", "add", "A1:B5", "--type", "pie", "--title", "원형", "--at", "H20", "--doc", "fresh")
    xl("save", "--as", "fresh.xlsx", "--doc", "fresh")
    theme = zipfile.ZipFile(os.path.join(ws, "fresh.xlsx")).read("xl/theme/theme1.xml").decode()
    accents = re.findall(r"<a:accent\d>\s*<a:srgbClr val=\"([0-9A-F]{6})\"", theme)
    check("new workbook gets the Office 2013-2022 theme", accents == ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
          and re.search(r"<a:majorFont>\s*<a:latin typeface=\"Calibri Light\"", theme) is not None, str(accents))
    c = charts_xml(os.path.join(ws, "fresh.xlsx"))
    bar, pie = c["막대"], c["원형"]
    line = next(v for k, v in c.items() if "lineChart" in v)
    check("no rounded corners, light chart-area border", attrs(bar, "roundedCorners") == ["0"] and "D9D9D9" in bar)
    check("columns: gap width 219, overlap -27", attrs(bar, "gapWidth") == ["219"] and attrs(bar, "overlap") == ["-27"])
    check("series colors are the theme accents in order", re.findall(r"<ser>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", bar, re.S)[:3] == ["4472C4", "ED7D31", "A5A5A5"],
          str(re.findall(r"<ser>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", bar, re.S)[:3]))
    check("text 9 pt / title 14 pt, gray 595959", set(attrs(bar, "defRPr", "sz")) == {"900", "1400"} and "595959" in bar, str(set(attrs(bar, "defRPr", "sz"))))
    check("value axis has light gridlines, category axis none", re.search(r"<valAx>.*<majorGridlines>.*D9D9D9.*</valAx>", bar, re.S) is not None
          and "<majorGridlines" not in re.search(r"<catAx>.*?</catAx>", bar, re.S).group(0))
    check("line: 2.25 pt, no markers", "w=\"28575\"" in line and attrs(line, "symbol") == ["none"])
    check("line without a title: autoTitleDeleted", attrs(line, "autoTitleDeleted") == ["1"])
    check("pie: a color per slice from the theme, white borders", attrs(pie, "srgbClr")[:3] == ["4472C4", "ED7D31", "A5A5A5"] or
          re.findall(r"<dPt>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", pie, re.S)[:3] == ["4472C4", "ED7D31", "A5A5A5"], str(re.findall(r"<dPt>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", pie, re.S)[:3]))

    # ---- 2) an existing workbook: its own theme, its own chart colors
    xl("open", "old.xlsx")
    d = json.loads(xl("describe", "--json", "--doc", "old.xlsx").stdout)
    ids = {ch["title"]: ch["id"] for ch in d["sheets"][0]["charts"]}
    xl("chart", "add", "A1:C4", "--title", "새차트", "--at", "E40", "--doc", "old.xlsx")
    xl("chart", "set", ids["colored"], "--title", "colored2", "--doc", "old.xlsx")
    xl("chart", "set", ids["slices"], "--title", "slices2", "--doc", "old.xlsx")
    xl("save", "--as", "old-out.xlsx", "--doc", "old.xlsx")
    c = charts_xml(os.path.join(ws, "old-out.xlsx"))
    sers = lambda s: re.findall(r"<ser>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", s, re.S)  # noqa: E731
    check("new chart in an Office 2007-theme workbook uses that theme (as Excel would)", sers(c["새차트"])[:2] == ["4F81BD", "C0504D"], str(sers(c["새차트"])[:2]))
    check("edited chart keeps its own series colors (incl. accent6 darker 25 %)", sers(c["colored2"])[:2] == ["C00000", "E46C0A"], str(sers(c["colored2"])[:2]))
    check("edited pie keeps its slice colors", re.findall(r"<dPt>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", c["slices2"], re.S)[:2] == ["112233", "C0504D"],
          str(re.findall(r"<dPt>.*?<a:srgbClr val=\"([0-9A-F]{6})\"", c["slices2"], re.S)[:2]))

    # ---- 3) review fixes: radar axes, color slots, pie slices, one-color pie, shade/tint
    from server.xlsx_io import _build_chart
    from server import colors as C

    def col(ser):
        gp = ser.graphicalProperties
        c = gp.solidFill or (gp.line.solidFill if gp.line is not None else None)
        return c.srgbClr if c is not None else None
    r = _build_chart({"type": "radar", "series": [{"val": "S!$B$2:$B$5", "cat": "S!$A$2:$A$5"}]}, None)
    check("radar: axes kept (delete=0) and styled", r.x_axis.delete is False and r.y_axis.delete is False and r.y_axis.txPr is not None and r.y_axis.majorGridlines is not None)
    cb = _build_chart({"type": "bar", "series": [{"val": "S!$B$2:$B$5", "name": "A"}, {"val": "S!$C$2:$C$5", "name": "B", "type": "line"}, {"val": "S!$D$2:$D$5", "name": "C"}]}, None)
    got = {s.tx.v: col(s) for ch in cb._charts for s in ch.series}
    check("combo bar+line: colors by series position, as the preview", got == {"A": "4472C4", "B": "ED7D31", "C": "A5A5A5"}, str(got))
    ce = _build_chart({"type": "bar", "series": [{"name": "X"}, {"val": "S!$B$2:$B$5", "name": "S2"}, {"val": "S!$C$2:$C$5", "name": "S3"}]}, None)
    got = {s.tx.v: col(s) for s in ce.series}
    check("a series without data still takes its color slot", got == {"S2": "ED7D31", "S3": "A5A5A5"}, str(got))
    p3 = _build_chart({"type": "pie", "series": [{"val": "S!$B$2:$B$4"}]}, None)
    p20 = _build_chart({"type": "pie", "series": [{"val": "S!$B$2:$B$21"}]}, None)
    check("pie: one slice color per value (3 -> 3, 20 -> 20)", len(p3.series[0].dPt) == 3 and len(p20.series[0].dPt) == 20)
    p1 = _build_chart({"type": "pie", "varyColors": False, "series": [{"val": "S!$B$2:$B$4", "color": "#C00000"}]}, None)
    check("one-color pie stays one color", p1.varyColors is False and col(p1.series[0]) == "C00000" and not p1.series[0].dPt)
    from openpyxl.drawing.colors import ColorChoice as CC, SchemeColor as SC
    check("shade/tint blend in linear RGB toward black/white", C.drawing_color(CC(schemeClr=SC(val="accent1", shade=50000)), C.DEFAULT_THEME) == "2F528F"
          and C.drawing_color(CC(schemeClr=SC(val="accent1", tint=40000)), C.DEFAULT_THEME) == "CFD5EA")
    # a one-color pie in a file is read as such
    from openpyxl.chart import PieChart as PC
    wb2 = Workbook()
    s2 = wb2.active
    for r_ in [["k", "v"], ["a", 1], ["b", 2]]:
        s2.append(r_)
    pc = PC()
    pc.varyColors = False
    pc.add_data(Reference(s2, min_col=2, min_row=1, max_row=3), titles_from_data=True)
    pc.series[0].graphicalProperties = GraphicalProperties(solidFill="C00000")
    s2.add_chart(pc, "D2")
    wb2.save(os.path.join(ws, "onecolor.xlsx"))
    from server.app import read_workbook
    m = read_workbook(os.path.join(ws, "onecolor.xlsx"))["sheets"][0]["charts"][0]
    check("reader keeps varyColors=0 and the series color", m.get("varyColors") is False and m["series"][0].get("color") == "#C00000", str(m.get("varyColors")))
finally:
    xl("stop", "--force")
    shutil.rmtree(ws, ignore_errors=True)

print("\nALL PASSED" if not fails else f"\n{fails} FAILED")
sys.exit(1 if fails else 0)
