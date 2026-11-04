"""xlsx <-> JSON workbook model, implemented with openpyxl.

Write strategy is *patching*: when the original file is available we load it again and
replace sheet contents, so parts the web UI does not model (conditional formats, data
validation, images, print settings, defined names, VBA in .xlsm …) survive the round trip.
"""
from __future__ import annotations

import datetime as dt
import io
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import (AreaChart, BarChart, DoughnutChart, LineChart, PieChart,
                            RadarChart, ScatterChart)
from openpyxl.chart.data_source import AxDataSource, NumRef, StrRef
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart import Series  # SeriesFactory
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, to_excel
from openpyxl.worksheet.filters import (AutoFilter, CustomFilter, CustomFilters, FilterColumn,
                                        Filters)
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.formula import ArrayFormula

from . import colors
from .formula_compat import add_prefixes, strip_prefixes
from .units import (DEFAULT_COL_PX, DEFAULT_ROW_PX, EMU_PER_PX, colw_to_px, pt_to_px,
                    px_to_colw, px_to_pt)

warnings.filterwarnings("ignore", module="openpyxl")

CHART_TYPES = {
    "barChart": "bar", "bar3DChart": "bar", "lineChart": "line", "line3DChart": "line",
    "pieChart": "pie", "pie3DChart": "pie", "ofPieChart": "pie", "doughnutChart": "doughnut",
    "areaChart": "area", "area3DChart": "area", "scatterChart": "scatter",
    "bubbleChart": "scatter", "radarChart": "radar",
}


# --------------------------------------------------------------------------- read

def _excel_value(v):
    if isinstance(v, dt.datetime):
        return to_excel(v)
    if isinstance(v, dt.date):
        return to_excel(dt.datetime(v.year, v.month, v.day))
    if isinstance(v, dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, dt.timedelta):
        return v.total_seconds() / 86400
    return v


def _side(side, theme):
    if side is None or not side.style:
        return None
    out = {"s": side.style}
    c = colors.resolve(side.color, theme)
    if c:
        out["c"] = c
    return out


def _style_dict(cell, theme, default_font):
    st = {}
    f = cell.font
    if f is not None:
        if f.b:
            st["b"] = True
        if f.i:
            st["i"] = True
        if f.u and f.u != "none":
            st["u"] = f.u if f.u in ("double",) else True
        if f.strike:
            st["st"] = True
        if f.sz and float(f.sz) != float(default_font[1]):
            st["fs"] = float(f.sz)
        if f.name and f.name != default_font[0]:
            st["ff"] = f.name
        c = colors.resolve(f.color, theme)
        if c and c != "#000000":
            st["fc"] = c
    fill = cell.fill
    if fill is not None:
        if getattr(fill, "patternType", None) and fill.patternType != "none":
            c = colors.resolve(fill.fgColor, theme) or colors.resolve(fill.bgColor, theme)
            if c:
                st["bg"] = c
        elif getattr(fill, "stop", None):
            c = colors.resolve(fill.stop[0].color, theme)
            if c:
                st["bg"] = c
    a = cell.alignment
    if a is not None:
        if a.horizontal and a.horizontal != "general":
            st["ha"] = a.horizontal
        if a.vertical and a.vertical != "bottom":
            st["va"] = a.vertical
        if a.wrap_text:
            st["wrap"] = True
        if a.indent:
            st["ind"] = int(a.indent)
        if a.text_rotation:
            st["rot"] = int(a.text_rotation)
    b = cell.border
    if b is not None:
        for key, side in (("bl", b.left), ("br", b.right), ("bt", b.top), ("bb", b.bottom)):
            s = _side(side, theme)
            if s:
                st[key] = s
    nf = cell.number_format
    if nf and nf != "General":
        st["nf"] = nf
    return st


def _rich_text(title):
    try:
        if title is None:
            return None
        if title.tx is not None and title.tx.rich is not None:
            out = []
            for p in title.tx.rich.p or []:
                out.append("".join((r.t or "") for r in (p.r or [])))
            return "\n".join(out) or None
        if title.tx is not None and title.tx.strRef is not None:
            return "=" + title.tx.strRef.f
    except Exception:
        pass
    return None


def _ref_of(ds):
    if ds is None:
        return None
    for attr in ("numRef", "strRef", "multiLvlStrRef"):
        r = getattr(ds, attr, None)
        if r is not None and getattr(r, "f", None):
            return r.f
    return None


class _Geom:
    """Pixel geometry of a worksheet, used to convert chart anchors."""

    def __init__(self, cols, rows, dcol, drow):
        self.cols, self.rows, self.dcol, self.drow = cols, rows, dcol, drow

    def cw(self, c):
        d = self.cols.get(c)
        return d["w"] if d and "w" in d else self.dcol

    def rh(self, r):
        d = self.rows.get(r)
        return d["h"] if d and "h" in d else self.drow

    def span_w(self, c1, c2):
        return sum(self.cw(c) for c in range(c1, c2))

    def span_h(self, r1, r2):
        return sum(self.rh(r) for r in range(r1, r2))

    def locate_x(self, x):
        c = 0
        while x >= self.cw(c) and c < 16383:
            x -= self.cw(c)
            c += 1
        return c, x

    def locate_y(self, y):
        r = 0
        while y >= self.rh(r) and r < 1048575:
            y -= self.rh(r)
            r += 1
        return r, y


def _anchor_model(anchor, geom: _Geom):
    E = EMU_PER_PX
    kind = type(anchor).__name__
    if kind == "OneCellAnchor":
        f = anchor._from
        return {"c": f.col, "r": f.row, "dx": round(f.colOff / E), "dy": round(f.rowOff / E),
                "w": round(anchor.ext.width / E), "h": round(anchor.ext.height / E)}
    if kind == "TwoCellAnchor":
        f, t = anchor._from, anchor.to
        x1 = geom.span_w(0, f.col) + f.colOff / E
        x2 = geom.span_w(0, t.col) + t.colOff / E
        y1 = geom.span_h(0, f.row) + f.rowOff / E
        y2 = geom.span_h(0, t.row) + t.rowOff / E
        return {"c": f.col, "r": f.row, "dx": round(f.colOff / E), "dy": round(f.rowOff / E),
                "w": max(40, round(x2 - x1)), "h": max(40, round(y2 - y1))}
    if kind == "AbsoluteAnchor":
        c, dx = geom.locate_x(anchor.pos.x / E)
        r, dy = geom.locate_y(anchor.pos.y / E)
        return {"c": c, "r": r, "dx": round(dx), "dy": round(dy),
                "w": round(anchor.ext.width / E), "h": round(anchor.ext.height / E)}
    return {"c": 4, "r": 1, "dx": 0, "dy": 0, "w": 480, "h": 288}


def _chart_model(chart, idx, geom, warn):
    subs = list(getattr(chart, "_charts", None) or [chart])
    main_type = CHART_TYPES.get(chart.tagname)
    if main_type is None:
        warn.append(f"지원하지 않는 차트 유형({chart.tagname})은 막대형으로 표시됩니다.")
        main_type = "bar"
    series = []
    for sub in subs:
        stype = CHART_TYPES.get(sub.tagname, main_type)
        for s in sub.series or []:
            name, name_ref = None, None
            if s.tx is not None:
                if s.tx.strRef is not None:
                    name_ref = s.tx.strRef.f
                else:
                    name = s.tx.v
            if stype == "scatter":
                cat, val = _ref_of(s.xVal), _ref_of(s.yVal)
            else:
                cat, val = _ref_of(s.cat), _ref_of(s.val)
            item = {"name": name, "nameRef": name_ref, "cat": cat, "val": val}
            if stype != main_type:
                item["type"] = stype
            series.append(item)
    m = {
        "type": main_type,
        "title": _rich_text(chart.title),
        "legend": chart.legend.position if chart.legend is not None else None,
        "series": series,
        "anchor": _anchor_model(chart.anchor, geom),
        "orig": idx,
    }
    if main_type == "bar":
        m["barDir"] = getattr(chart, "barDir", "col") or "col"
        m["grouping"] = getattr(chart, "grouping", "clustered") or "clustered"
    elif main_type in ("line", "area"):
        m["grouping"] = getattr(chart, "grouping", "standard") or "standard"
    try:
        m["xTitle"] = _rich_text(chart.x_axis.title)
        m["yTitle"] = _rich_text(chart.y_axis.title)
    except Exception:
        pass
    return m


def _filter_model(af):
    if af is None or not af.ref:
        return None
    out = {"ref": af.ref, "cols": {}}
    for fc in af.filterColumn or []:
        entry = {}
        if fc.filters is not None:
            entry["values"] = [str(x) for x in (fc.filters.filter or [])]
            if fc.filters.blank:
                entry["blank"] = True
        if fc.customFilters is not None:
            entry["custom"] = [{"op": cf.operator or "equal", "val": cf.val}
                               for cf in (fc.customFilters.customFilter or [])]
            entry["and"] = bool(fc.customFilters._and)
        if entry:
            out["cols"][str(fc.colId)] = entry
    return out


def read_xlsx(data: bytes, keep_vba: bool = False) -> dict:
    wb = load_workbook(io.BytesIO(data), data_only=False, keep_vba=keep_vba)
    try:
        wbv = load_workbook(io.BytesIO(data), data_only=True, read_only=False)
    except Exception:
        wbv = None
    theme = colors.parse_theme_colors(getattr(wb, "loaded_theme", None))
    try:
        df = wb._fonts[0]
        default_font = (df.name or "Calibri", float(df.sz or 11))
    except Exception:
        default_font = ("Calibri", 11.0)

    styles, style_index = [], {}
    warn: list[str] = []
    sheets = []
    for ws in wb.worksheets:
        wsv = wbv[ws.title] if wbv is not None else None
        sf = ws.sheet_format
        dcol = colw_to_px(sf.defaultColWidth) if sf.defaultColWidth else (
            colw_to_px((sf.baseColWidth or 8) + 0.43))
        drow = pt_to_px(sf.defaultRowHeight) if sf.defaultRowHeight else DEFAULT_ROW_PX

        cols = {}
        for key, d in ws.column_dimensions.items():
            lo, hi = d.min or 0, d.max or 0
            if not lo:
                continue
            hi = min(hi or lo, lo + 2000)  # guard against A:XFD style ranges
            for c in range(lo - 1, hi):
                e = {}
                if d.width and d.customWidth is not False:
                    e["w"] = colw_to_px(d.width)
                if d.hidden:
                    e["hidden"] = True
                if e:
                    cols[c] = e
        rows = {}
        for r, d in ws.row_dimensions.items():
            e = {}
            if d.ht is not None and d.customHeight:
                e["h"] = pt_to_px(d.ht)
            elif d.ht is not None:
                e["h"] = pt_to_px(d.ht)
            if d.hidden:
                e["hidden"] = True
            if e:
                rows[r - 1] = e

        cells = []
        for (r, c), cell in ws._cells.items():
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            item = {}
            if isinstance(v, ArrayFormula):
                item["f"] = strip_prefixes(v.text if v.text.startswith("=") else "=" + v.text)
                item["a"] = v.ref
                v = wsv._cells[(r, c)].value if wsv is not None and (r, c) in wsv._cells else None
            elif cell.data_type == "f" and isinstance(v, str):
                item["f"] = strip_prefixes(v)
                v = wsv._cells[(r, c)].value if wsv is not None and (r, c) in wsv._cells else None
            elif not isinstance(v, (str, int, float, bool, type(None), dt.date, dt.time, dt.timedelta)):
                v = None  # DataTableFormula etc.
            v = _excel_value(v)
            if v is not None:
                item["v"] = v
            if cell.data_type == "e" or (isinstance(v, str) and v.startswith("#") and "f" in item
                                         and v in ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!",
                                                   "#NAME?", "#NUM!", "#N/A")):
                item["t"] = "e"
            if cell.has_style:
                key = tuple(cell._style)
                si = style_index.get(key)
                if si is None:
                    sd = _style_dict(cell, theme, default_font)
                    si = len(styles) if sd else -1
                    if sd:
                        styles.append(sd)
                    style_index[key] = si
                if si >= 0:
                    item["s"] = si
            if cell.hyperlink is not None and (cell.hyperlink.target or cell.hyperlink.location):
                item["h"] = cell.hyperlink.target or ("#" + cell.hyperlink.location)
            if cell.comment is not None:
                item["cm"] = cell.comment.text
            if item:
                cells.append([r - 1, c - 1, item])

        geom = _Geom(cols, rows, dcol, drow)
        charts = []
        for i, ch in enumerate(getattr(ws, "_charts", [])):
            try:
                charts.append(_chart_model(ch, i, geom, warn))
            except Exception as exc:  # pragma: no cover
                warn.append(f"차트를 읽지 못했습니다: {exc}")
        if getattr(ws, "_images", None):
            warn.append(f"'{ws.title}' 시트의 이미지는 화면에 표시되지 않지만 저장 시 보존됩니다.")
        if ws.conditional_formatting:
            warn.append(f"'{ws.title}' 시트의 조건부 서식은 화면에 표시되지 않지만 저장 시 보존됩니다.")

        tab = ws.sheet_properties.tabColor
        sheets.append({
            "name": ws.title,
            "origName": ws.title,
            "state": ws.sheet_state,
            "tabColor": colors.resolve(tab, theme) if tab is not None else None,
            "defaultColW": dcol or DEFAULT_COL_PX,
            "defaultRowH": drow or DEFAULT_ROW_PX,
            "cols": {str(k): v for k, v in cols.items()},
            "rows": {str(k): v for k, v in rows.items()},
            "cells": cells,
            "merges": [str(m) for m in ws.merged_cells.ranges],
            "freeze": ws.freeze_panes,
            "autoFilter": _filter_model(ws.auto_filter),
            "charts": charts,
            "showGrid": ws.sheet_view.showGridLines is not False,
        })

    names = {}
    try:
        for name, dn in wb.defined_names.items():
            if dn.attr_text:
                names[name] = strip_prefixes(dn.attr_text)
    except Exception:
        pass
    if wb.chartsheets:
        warn.append("차트 시트는 화면에 표시되지 않지만 저장 시 보존됩니다.")
    active = wb.worksheets.index(wb.active) if wb.active in wb.worksheets else 0
    return {"sheets": sheets, "styles": styles, "activeSheet": active,
            "definedNames": names, "warnings": warn,
            "defaultFont": {"name": default_font[0], "size": default_font[1]}}


# --------------------------------------------------------------------------- write

def _color(hex_):
    return "FF" + hex_.lstrip("#").upper() if hex_ else None


class _StyleCache:
    def __init__(self, default_font):
        self.default_font = default_font
        self.cache = {}

    def get(self, sd: dict):
        key = repr(sorted(sd.items(), key=lambda kv: kv[0]))
        hit = self.cache.get(key)
        if hit is not None:
            return hit
        u = sd.get("u")
        font = Font(name=sd.get("ff", self.default_font[0]), sz=sd.get("fs", self.default_font[1]),
                    b=bool(sd.get("b")), i=bool(sd.get("i")),
                    u=("double" if u == "double" else "single") if u else None,
                    strike=bool(sd.get("st")),
                    color=_color(sd.get("fc")))
        fill = PatternFill("solid", fgColor=_color(sd["bg"])) if sd.get("bg") else PatternFill()

        def side(k):
            s = sd.get(k)
            return Side(style=s["s"], color=_color(s.get("c", "#000000"))) if s else Side()

        border = Border(left=side("bl"), right=side("br"), top=side("bt"), bottom=side("bb"))
        align = Alignment(horizontal=sd.get("ha"), vertical=sd.get("va"),
                          wrap_text=bool(sd.get("wrap")) or None,
                          indent=sd.get("ind", 0), text_rotation=sd.get("rot", 0))
        hit = (font, fill, border, align, sd.get("nf", "General"))
        self.cache[key] = hit
        return hit


def _build_chart(m: dict):
    kinds = {"bar": BarChart, "line": LineChart, "pie": PieChart, "doughnut": DoughnutChart,
             "area": AreaChart, "scatter": ScatterChart, "radar": RadarChart}
    main_t = m.get("type", "bar")
    groups: dict[str, list] = {}
    for s in m.get("series", []):
        groups.setdefault(s.get("type") or main_t, []).append(s)
    if main_t not in groups:
        groups = {main_t: [], **groups}

    counter = [0]

    def make(t, series):
        ch = kinds.get(t, BarChart)()
        if t == "bar":
            ch.type = m.get("barDir", "col")
            g = m.get("grouping", "clustered")
            ch.grouping = g if g in ("clustered", "stacked", "percentStacked") else "clustered"
            if ch.grouping != "clustered":
                ch.overlap = 100
        elif t in ("line", "area"):
            g = m.get("grouping", "standard")
            ch.grouping = g if g in ("standard", "stacked", "percentStacked") else "standard"
        if t == "scatter":
            ch.style = 13
        for s in series:
            if not s.get("val"):
                continue
            if t == "scatter":
                ser = Series(s["val"], xvalues=s.get("cat") or None)
                ser.marker = Marker(symbol="circle")
                ser.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
            else:
                ser = Series(s["val"])
                if s.get("cat"):
                    ser.cat = AxDataSource(numRef=NumRef(f=s["cat"]))
            ser.idx = ser.order = counter[0]
            counter[0] += 1
            if s.get("nameRef"):
                ser.tx = SeriesLabel(strRef=StrRef(f=s["nameRef"]))
            elif s.get("name"):
                ser.tx = SeriesLabel(v=str(s["name"]))
            ch.series.append(ser)
        for ax_name in ("x_axis", "y_axis"):
            ax = getattr(ch, ax_name, None)
            if ax is not None:
                ax.delete = False  # openpyxl>=3.1 hides axes otherwise in Excel 365
        return ch

    items = list(groups.items())
    chart = make(*items[0])
    for t, series in items[1:]:
        chart += make(t, series)
    if m.get("title"):
        chart.title = m["title"]
    if m.get("legend"):
        chart.legend = Legend()
        chart.legend.position = m["legend"]
    else:
        chart.legend = None
    try:
        if m.get("xTitle"):
            chart.x_axis.title = m["xTitle"]
        if m.get("yTitle"):
            chart.y_axis.title = m["yTitle"]
    except Exception:
        pass
    return chart


def _anchor(a: dict):
    E = EMU_PER_PX
    return OneCellAnchor(
        _from=AnchorMarker(col=int(a.get("c", 0)), colOff=int(a.get("dx", 0) * E),
                           row=int(a.get("r", 0)), rowOff=int(a.get("dy", 0) * E)),
        ext=XDRPositiveSize2D(cx=int(a.get("w", 480) * E), cy=int(a.get("h", 288) * E)))


def _write_filter(ws, fm):
    if not fm or not fm.get("ref"):
        ws.auto_filter = AutoFilter()
        return
    af = AutoFilter(ref=fm["ref"])
    for col_id, e in sorted(fm.get("cols", {}).items(), key=lambda kv: int(kv[0])):
        fc = FilterColumn(colId=int(col_id))
        if e.get("values") is not None:
            fc.filters = Filters(filter=[str(v) for v in e["values"]], blank=bool(e.get("blank")))
        if e.get("custom"):
            fc.customFilters = CustomFilters(
                _and=bool(e.get("and")),
                customFilter=[CustomFilter(operator=c.get("op", "equal"), val=str(c.get("val", "")))
                              for c in e["custom"]])
        af.filterColumn.append(fc)
    ws.auto_filter = af


def write_xlsx(model: dict, base: bytes | None = None, keep_vba: bool = False) -> bytes:
    if base:
        wb = load_workbook(io.BytesIO(base), keep_vba=keep_vba)
        wb.epoch = CALENDAR_WINDOWS_1900  # the web model always uses 1900-based serials
    else:
        wb = Workbook()
        wb.remove(wb.active)
        dfont = model.get("defaultFont") or {}
        if dfont.get("name"):
            from openpyxl.utils.indexed_list import IndexedList
            f0 = Font(name=dfont["name"], sz=float(dfont.get("size") or 11))
            wb._fonts = IndexedList([f0])
            try:
                wb._named_styles["Normal"].font = f0
            except Exception:
                pass
    try:
        df = wb._fonts[0]
        default_font = (df.name or "Calibri", float(df.sz or 11))
    except Exception:
        default_font = ("Calibri", 11.0)
    cache = _StyleCache(default_font)
    styles = model.get("styles", [])

    existing = {ws.title: ws for ws in wb.worksheets}
    reused = {}
    for sm in model["sheets"]:
        on = sm.get("origName")
        if on and on in existing and on not in reused:
            reused[on] = existing[on]
    for i, ws in enumerate(reused.values()):  # avoid rename collisions
        ws.title = f"__tmp_rename_{i}__"
    for name, ws in existing.items():
        if name not in reused:
            wb.remove(ws)

    ordered = []
    for sm in model["sheets"]:
        on = sm.get("origName")
        ws = reused.pop(on, None) if on else None
        if ws is None:
            ws = wb.create_sheet()
        ws.title = sm["name"]
        ordered.append((ws, sm))
    wb._sheets = [ws for ws, _ in ordered] + list(wb.chartsheets)

    for ws, sm in ordered:
        orig_charts = list(getattr(ws, "_charts", []))
        ws._cells.clear()
        ws._hyperlinks = []
        ws.merged_cells = MultiCellRange()
        ws.column_dimensions.clear()
        ws.row_dimensions.clear()

        dcol = sm.get("defaultColW") or DEFAULT_COL_PX
        drow = sm.get("defaultRowH") or DEFAULT_ROW_PX
        if dcol != DEFAULT_COL_PX:
            ws.sheet_format.defaultColWidth = px_to_colw(dcol)
        if drow != DEFAULT_ROW_PX:
            ws.sheet_format.defaultRowHeight = px_to_pt(drow)
            ws.sheet_format.customHeight = True

        for key, e in sm.get("cols", {}).items():
            letter = get_column_letter(int(key) + 1)
            d = ws.column_dimensions[letter]
            if "w" in e:
                d.width = px_to_colw(e["w"])
            if e.get("hidden"):
                d.hidden = True
        for key, e in sm.get("rows", {}).items():
            d = ws.row_dimensions[int(key) + 1]
            if "h" in e:
                d.height = px_to_pt(e["h"])
            if e.get("hidden"):
                d.hidden = True

        for r, c, item in sm.get("cells", []):
            cell = ws.cell(row=r + 1, column=c + 1)
            f = item.get("f")
            if f:
                text = add_prefixes(f if f.startswith("=") else "=" + f)
                cell.value = ArrayFormula(item["a"], text) if item.get("a") else text
            else:
                v = item.get("v")
                if isinstance(v, str) and item.get("t") != "e":
                    cell.value = v
                    cell.data_type = "s"
                else:
                    cell.value = v
            si = item.get("s")
            if isinstance(si, int) and 0 <= si < len(styles) and styles[si]:
                font, fill, border, align, nf = cache.get(styles[si])
                cell.font, cell.fill, cell.border, cell.alignment = font, fill, border, align
                cell.number_format = nf
            if item.get("h"):
                cell.hyperlink = item["h"]
            if item.get("cm"):
                cell.comment = Comment(item["cm"], "web-excel")

        for ref in sm.get("merges", []):
            ws.merge_cells(ref)
        ws.freeze_panes = sm.get("freeze") or None
        _write_filter(ws, sm.get("autoFilter"))
        ws.sheet_state = sm.get("state", "visible")
        ws.sheet_properties.tabColor = _color(sm["tabColor"])[2:] if sm.get("tabColor") else None
        ws.sheet_view.showGridLines = sm.get("showGrid", True)

        new_charts = []
        for cm in sm.get("charts", []):
            o = cm.get("orig")
            if isinstance(o, int) and not cm.get("dirty") and 0 <= o < len(orig_charts):
                ch = orig_charts[o]
            else:
                ch = _build_chart(cm)
            ch.anchor = _anchor(cm.get("anchor", {}))
            new_charts.append(ch)
        ws._charts = new_charts

    active = model.get("activeSheet", 0)
    if wb.worksheets:
        wb.active = max(0, min(active, len(wb.worksheets) - 1))
        for ws in wb.worksheets:
            ws.sheet_view.tabSelected = ws is wb.active
    wb.calculation.fullCalcOnLoad = True  # no cached values are written; let Excel recalc

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
