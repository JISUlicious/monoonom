"""One chart style for the preview and for saved files: Excel 2016+ default chart ("Style 1").

Charts web-excel creates or edits are written with these values explicitly (colors as RGB, text
sizes and colors, gridlines, chart area), so Excel draws them the same whatever theme the workbook
has, and the preview (assets/web/js/charts.js, CHART_STYLE) uses the same numbers. Untouched charts
from the original file keep their own XML.

Values: Excel 2013-2022 default chart formatting (text tx1 65% = 595959, lines bg1 85% = D9D9D9,
title 14 pt, other text 9 pt, column gap width 219 % / overlap -27 %, line 2.25 pt without markers),
DrawingML units: sizes in 1/100 pt, widths in EMU (12700 per pt). ECMA-376 Part 1 §21.2 (charts),
§20.1.2.3 (color transforms lumMod/lumOff).
"""
from __future__ import annotations

from openpyxl.chart.axis import ChartLines
from openpyxl.chart.layout import Layout  # noqa: F401  (kept for callers that position the plot area)
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun

from . import colors

TEXT = "595959"        # tx1, lumMod 65 % / lumOff 35 %
LINE = "D9D9D9"        # bg1, lumMod 85 % (gridlines, category axis, chart area border)
TITLE_SZ = 1400        # 14 pt
TEXT_SZ = 900          # 9 pt
THIN = 9525            # 0.75 pt
SERIES_LINE = 28575    # 2.25 pt
PIE_BORDER = 12700     # 1 pt, white
BAR_GAP, BAR_OVERLAP = 219, -27
STACKED_GAP, STACKED_OVERLAP = 150, 100
SCATTER_MARKER = 5


def palette(theme: list[str] | None, n: int = 18) -> list[str]:
    """Series colors in Excel's order: accent1-6, then the same darker (lumMod 60 %), then lighter
    (lumMod 80 % + lumOff 20 %). theme: the 12 theme colors (lt1, dk1, lt2, dk2, accent1-6, ...)."""
    acc = list((theme or colors.DEFAULT_THEME)[4:10])
    out = acc + [colors.lum(a, 60000, 0) for a in acc] + [colors.lum(a, 80000, 20000) for a in acc]
    return [out[i % len(out)] for i in range(n)]


def _text(sz: int, bold: bool = False) -> CharacterProperties:
    return CharacterProperties(sz=sz, b=bold, solidFill=TEXT)


def rich(sz: int = TEXT_SZ) -> RichText:
    """txPr for axes and the legend."""
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=_text(sz)), endParaRPr=_text(sz))])


def title(text: str) -> Title:
    cp = _text(TITLE_SZ)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def axis_title(text: str) -> Title:
    cp = _text(1000)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def _line(color: str | None, w: int = THIN) -> LineProperties:
    return LineProperties(solidFill=color, w=w) if color else LineProperties(noFill=True)


def series_props(kind: str, color: str) -> GraphicalProperties:
    if kind in ("line", "radar"):
        return GraphicalProperties(ln=LineProperties(solidFill=color, w=SERIES_LINE, cap="rnd"))
    if kind == "scatter":
        return GraphicalProperties(ln=LineProperties(noFill=True))
    return GraphicalProperties(solidFill=color, ln=LineProperties(noFill=True))


def chart_area() -> GraphicalProperties:
    return GraphicalProperties(solidFill="FFFFFF", ln=_line(LINE))


def style_axes(chart) -> None:
    """Category axis: light line, no gridlines; value axis: no line, light major gridlines."""
    x, y = getattr(chart, "x_axis", None), getattr(chart, "y_axis", None)
    for ax in (x, y):
        if ax is None:
            continue
        ax.txPr = rich()
        ax.majorTickMark = "none"
        ax.minorTickMark = "none"
        ax.delete = False  # openpyxl >= 3.1 hides axes otherwise in Excel 365
    scatter = getattr(chart, "tagname", "") == "scatterChart"
    if x is not None:
        x.spPr = GraphicalProperties(ln=_line(LINE))
        x.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=_line(LINE))) if scatter else None
    if y is not None:
        y.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
        y.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=_line(LINE)))


def style_radar_axes(chart) -> None:
    """Radar: keep the axes (openpyxl leaves delete unset, which Excel 365 reads as hidden), 9 pt
    gray labels and light gridlines, as the preview draws them."""
    for ax in (getattr(chart, "x_axis", None), getattr(chart, "y_axis", None)):
        if ax is None:
            continue
        ax.delete = False
        ax.txPr = rich()
        ax.majorTickMark = "none"
        ax.minorTickMark = "none"
        ax.spPr = GraphicalProperties(ln=_line(LINE))
    if getattr(chart, "y_axis", None) is not None:
        chart.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=_line(LINE)))


def marker_none() -> Marker:
    return Marker(symbol="none")


def marker_circle(color: str) -> Marker:
    return Marker(symbol="circle", size=SCATTER_MARKER,
                  spPr=GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color, w=THIN)))


def pie_border() -> LineProperties:
    return LineProperties(solidFill="FFFFFF", w=PIE_BORDER)


def preview_spec() -> dict:
    """The same numbers for the browser (sent with the workbook so both stay in step)."""
    return {"text": "#" + TEXT, "line": "#" + LINE, "titlePt": TITLE_SZ / 100, "textPt": TEXT_SZ / 100,
            "seriesLinePt": SERIES_LINE / 12700, "gap": BAR_GAP, "overlap": BAR_OVERLAP,
            "stackedGap": STACKED_GAP, "scatterMarker": SCATTER_MARKER, "pieBorderPt": PIE_BORDER / 12700}
