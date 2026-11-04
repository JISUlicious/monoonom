"""Generate sample workbooks into ./samples (xlsx with styles/formulas/chart/filter).

samples/sample.xls is a committed fixture: web-excel no longer bundles an .xls writer (xlwt),
and make_xls() copies it.
"""
import datetime as dt
import os
import shutil
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "samples")


def make_xlsx(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "매출"
    head = ["월", "제품A", "제품B", "제품C", "합계", "날짜"]
    ws.append(head)
    for i, m in enumerate(["1월", "2월", "3월", "4월", "5월", "6월"], start=2):
        ws.append([m, 100 + i * 13, 80 + (i * 37) % 50, 60 + (i * 17) % 40,
                   f"=SUM(B{i}:D{i})", dt.date(2026, i - 1, 1)])
        ws.cell(i, 6).number_format = "yyyy-mm-dd"
    ws["A8"] = "합계"
    for col in "BCDE":
        ws[f"{col}8"] = f"=SUM({col}2:{col}7)"
    ws["A9"] = "평균"
    ws["B9"] = "=AVERAGE(B2:B7)"
    ws["B9"].number_format = "#,##0.00"
    ws["A10"] = "최대 제품A 월"
    ws["B10"] = '=INDEX(A2:A7,MATCH(MAX(B2:B7),B2:B7,0))'
    ws["A11"] = "조건"
    ws["B11"] = '=IF(B8>700,"목표달성","미달")'
    ws["A12"] = "XLOOKUP"
    ws["B12"] = '=XLOOKUP("3월",A2:A7,C2:C7)'
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    thin = Side(style="thin", color="999999")
    for c in ws[1]:
        c.font = Font(b=True, color="FFFFFF")
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center")
        c.border = Border(bottom=thin)
    for row in ws.iter_rows(min_row=8, max_row=8):
        for c in row:
            c.font = Font(b=True)
            c.border = Border(top=Side(style="thin"), bottom=Side(style="double"))
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["F"].width = 12
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:F7"
    ch = BarChart()
    ch.title = "월별 매출"
    ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=1, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=5, min_row=1, max_row=7), titles_from_data=True)
    ch += ln
    ws.add_chart(ch, "H2")

    ws2 = wb.create_sheet("비중")
    ws2.append(["제품", "점유율"])
    for n, v in (("A", 0.45), ("B", 0.35), ("C", 0.2)):
        ws2.append([n, v])
    for r in range(2, 5):
        ws2.cell(r, 2).number_format = "0%"
    ws2["D1"] = "=매출!E8"
    ws2.merge_cells("A6:B6")
    ws2["A6"] = "병합된 셀"
    ws2["A6"].alignment = Alignment(horizontal="center")
    pie = PieChart()
    pie.title = "점유율"
    pie.add_data(Reference(ws2, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    pie.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=4))
    ws2.add_chart(pie, "D3")
    wb.save(path)


def make_offset_xlsx(path):
    """A table that does not start at A1: title, blank rows, header at C5, totals row under the data."""
    wb = Workbook()
    ws = wb.active
    ws.title = "지점"
    ws["B2"] = "2026 지점별 매출"
    ws.merge_cells("B2:F2")
    rows = [["지점", "1월", "2월", "합계"], ["서울", 120, 130], ["부산", 90, 80], ["대구", 150, 60],
            ["서울", 120, 130], ["광주", 70, 95], ["대전", 110, 140], ["울산", 60, 50]]
    for r, row in enumerate(rows, start=5):
        for c, v in enumerate(row, start=3):
            ws.cell(r, c, v)
        if r > 5:
            ws.cell(r, 6, f"=D{r}+E{r}")
    ws["C13"] = "합계"
    for col in "DEF":
        ws[f"{col}13"] = f"=SUM({col}6:{col}12)"
    ws["H5"] = "메모: 단위 만원"
    wb.save(path)


def make_xls(path):
    """Legacy .xls fixture (sheet 데이터: 이름/수량/단가/금액, D2:D4 = B*C formulas)."""
    if os.path.abspath(path) != os.path.abspath(os.path.join(OUT, "sample.xls")):
        shutil.copyfile(os.path.join(OUT, "sample.xls"), path)

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else OUT
    os.makedirs(out, exist_ok=True)
    make_xlsx(os.path.join(out, "sample.xlsx"))
    make_xls(os.path.join(out, "sample.xls"))
    make_offset_xlsx(os.path.join(out, "offset.xlsx"))
    print("written to", out)
