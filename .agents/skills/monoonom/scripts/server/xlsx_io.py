"""xlsx <-> JSON workbook model, implemented with openpyxl.

Write strategy is *patching*: when the original file is available we load it again and
replace sheet contents, so parts the web UI does not model (conditional formats, data
validation, images, print settings, defined names, VBA in .xlsm …) survive the round trip.
"""
from __future__ import annotations

import datetime as dt
import io
import re
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import (AreaChart, BarChart, DoughnutChart, LineChart, PieChart,
                            RadarChart, ScatterChart)
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
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

from . import a1 as A
from . import colors
from . import pivots as PV
from .formula_compat import add_prefixes, prefix_map, strip_prefixes, unsupported_functions
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


def _lit_of(ds):
    """A series' own list of values (numLit / strLit: no cell range), in point order, or None."""
    if ds is None:
        return None
    for attr, num in (("numLit", True), ("strLit", False)):
        d = getattr(ds, attr, None)
        if d is not None and d.pt:
            n = max(d.ptCount or 0, max(p.idx for p in d.pt) + 1)
            out = [None] * n
            for p in d.pt:
                try:
                    out[p.idx] = float(p.v) if num else p.v
                except (TypeError, ValueError):
                    out[p.idx] = None
            return out
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


def _series_color(s, stype, theme):
    """Explicit color of a series in the file (fill for bars/areas/slices, line for lines,
    marker for scatter), or None when Excel picks it from the theme by position."""
    gp = getattr(s, "graphicalProperties", None)
    if stype in ("line", "radar"):
        ln = getattr(gp, "line", None) if gp is not None else None
        c = colors.drawing_color(getattr(ln, "solidFill", None), theme) if ln is not None else None
    elif stype == "scatter":
        mk = getattr(s, "marker", None)
        mgp = getattr(mk, "graphicalProperties", None) if mk is not None else None
        c = colors.drawing_color(getattr(mgp, "solidFill", None), theme) if mgp is not None else None
    else:
        c = colors.drawing_color(getattr(gp, "solidFill", None), theme) if gp is not None else None
    return "#" + c if c else None


def _chart_model(chart, idx, geom, warn, theme=None):
    subs = list(getattr(chart, "_charts", None) or [chart])
    main_type = CHART_TYPES.get(chart.tagname)
    if main_type is None:
        warn.append(f"Unsupported chart type ({chart.tagname}) is shown as a bar chart.")
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
            xs, ys = (s.xVal, s.yVal) if stype == "scatter" else (s.cat, s.val)
            if not cat and _lit_of(xs):
                item["catLit"] = _lit_of(xs)
            if not val and _lit_of(ys):
                item["valLit"] = _lit_of(ys)
            color = _series_color(s, stype, theme)
            if color:
                item["color"] = color
            if stype in ("pie", "doughnut") and getattr(s, "dPt", None):
                pts = []
                for dp in s.dPt:
                    c = colors.drawing_color(getattr(getattr(dp, "graphicalProperties", None), "solidFill", None), theme)
                    while len(pts) <= dp.idx:
                        pts.append(None)
                    pts[dp.idx] = "#" + c if c else None
                if any(pts):
                    item["pointColors"] = pts
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
    vary = getattr(chart, "varyColors", None)
    m["varyColors"] = True if vary is None else bool(vary)  # Excel reads a missing element as on
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


# ------------------------------------------------------------------ pivot tables (step 1: safeguards)
# A pivot table's output is ordinary cells in the sheet; its definition (location, source range, fields)
# and its cache stay in the file (openpyxl keeps them). The model describes each pivot so the app can
# protect its area, move it with row/column inserts, and ask Excel to refresh it when its source changed.

def _pivot_source(p, sheet_name):
    ws_src = getattr(getattr(p.cache, "cacheSource", None), "worksheetSource", None)
    if ws_src is None:
        return None
    if ws_src.name:  # a defined name or table
        return ws_src.name
    return f"{A.quote_sheet(ws_src.sheet or sheet_name)}!{ws_src.ref}"


def _shared_values(cf):
    """A cache field's shared items as plain values (text, number, bool, date as Excel serial)."""
    out = []
    for x in getattr(getattr(cf, "sharedItems", None), "_fields", None) or []:
        v = getattr(x, "v", None)
        if isinstance(v, (dt.date, dt.datetime)):
            v = _excel_value(v)
        out.append(v)
    return out


def _pivot_layout(p, names):
    """What refreshing in place needs (step 2), with item indices resolved to values:
    rowFields/colFields (field names, "Σ" for the data-field position), rowItems/colItems (Excel's
    list of what each output row/column shows: t data|default|grand|blank, r = leading fields kept
    from the previous item, x = item values, i = data field), hidden items per field (filters),
    and the data area's offset in the output (firstDataRow/Col)."""
    shared = [_shared_values(cf) for cf in (p.cache.cacheFields or [])]
    items = []  # pivot field k: item index -> value
    for k, pf in enumerate(p.pivotFields or []):
        vals = shared[k] if k < len(shared) else []
        items.append([vals[it.x] if it.x is not None and it.x < len(vals) else None for it in (pf.items or []) if it.t == "data"])
    hidden = {}
    for k, pf in enumerate(p.pivotFields or []):
        vals = shared[k] if k < len(shared) else []
        h = [vals[it.x] for it in (pf.items or []) if it.t == "data" and it.h and it.x is not None and it.x < len(vals)]
        if h:
            hidden[names[k]] = h
    for pg in p.pageFields or []:  # a report filter showing one item hides the others
        if pg.item is not None and 0 <= pg.fld < len(items) and pg.item < len(items[pg.fld]):
            keep = items[pg.fld][pg.item]
            hidden[names[pg.fld]] = [v for v in items[pg.fld] if v != keep]

    def axis(fields):
        return [("Σ" if f.x == -2 else names[f.x]) for f in (fields or []) if f.x == -2 or 0 <= f.x < len(names)]

    def resolve(row_items, fields):
        out = []
        for it in row_items or []:
            xs, d = [], None
            for j, x in enumerate(it.x or []):
                pos = (it.r or 0) + j
                f = fields[pos] if pos < len(fields) else None
                if f is None or f.x == -2:
                    xs.append(None)
                    if f is not None:
                        d = x.v or 0  # the data field shown here
                else:
                    lst = items[f.x] if f.x < len(items) else []
                    xs.append(lst[x.v or 0] if (x.v or 0) < len(lst) else None)
            if it.t == "grand":  # x carries no item for a grand total
                xs = [None] * len(xs)
            out.append({"t": it.t or "data", "r": it.r or 0, "i": it.i or 0, "x": xs, **({"d": d} if d is not None else {})})
        return out

    L = p.location
    return {"rowFields": axis(p.rowFields), "colFields": axis(p.colFields),
            "rowItems": resolve(p.rowItems, list(p.rowFields or [])) if p.rowItems else None,
            "colItems": resolve(p.colItems, list(p.colFields or [])) if p.colItems else None,
            "hidden": hidden, "firstDataRow": L.firstDataRow or 0, "firstDataCol": L.firstDataCol or 0,
            "firstHeaderRow": L.firstHeaderRow or 0, "items": {names[k]: v for k, v in enumerate(items) if k < len(names) and v},
            "fields": names}


def _pivot_model(p, idx, sheet_name):
    names = [f.name for f in (p.cache.cacheFields or [])] if p.cache is not None else []
    nm = lambda i: names[i] if isinstance(i, int) and 0 <= i < len(names) else None  # noqa: E731 (-2 = "Values")
    try:
        layout = _pivot_layout(p, names)
    except Exception:  # pragma: no cover - a pivot we cannot describe is still protected
        layout = None
    return {
        "layout": layout,
        "idx": idx, "name": p.name, "ref": p.location.ref,
        "source": _pivot_source(p, sheet_name),
        "rows": [nm(f.x) for f in (p.rowFields or []) if nm(f.x)],
        "cols": [nm(f.x) for f in (p.colFields or []) if nm(f.x)],
        "values": [{"name": d.name, "field": nm(d.fld), "fn": d.subtotal or "sum"} for d in (p.dataFields or [])],
        "filters": [nm(f.fld) for f in (p.pageFields or []) if nm(f.fld)],
    }


def pivot_source_sig(sheets, source):
    """Fingerprint of the source range's values and formulas (model sheets with cells as [r, c, item])."""
    import hashlib
    if not source or "!" not in source:
        return None
    sheet, ref = A.split_sheet(source)
    sm = next((s for s in sheets if s["name"].lower() == (sheet or "").lower()), None)
    if sm is None:
        return None
    try:
        r1, c1, r2, c2 = A.parse_range(ref)
    except ValueError:
        return None
    pts = sorted((r, c, item.get("v"), item.get("f")) for r, c, item in sm.get("cells", []) if r1 <= r <= r2 and c1 <= c <= c2
                 and (item.get("v") is not None or item.get("f")))
    return hashlib.sha1(repr(pts).encode("utf-8")).hexdigest()


def _write_pivots(ws, sm, sheets):
    """Keep each pivot's definition in step with the model: location and source after row/column
    inserts, sheet renames; refreshOnLoad when the source data changed (Excel then rebuilds the
    pivot when the file opens, so it never shows stale numbers)."""
    pivots = getattr(ws, "_pivots", None) or []
    for pm in sm.get("pivots") or []:
        if pm.get("create"):  # made in the app: write its definition (once; the save assigns its idx)
            pm["savedIdx"] = _create_pivot(ws.parent, ws, pm)
            continue
        i = pm.get("idx")
        if not isinstance(i, int) or not 0 <= i < len(pivots):
            continue
        p = pivots[i]
        if pm.get("ref") and p.location.ref != pm["ref"]:
            p.location.ref = pm["ref"]
        src = getattr(getattr(p.cache, "cacheSource", None), "worksheetSource", None)
        changed = pm.get("srcSig") is not None and pivot_source_sig(sheets, pm.get("source")) != pm["srcSig"]
        if src is not None and not src.name and pm.get("source") and "!" in pm["source"]:
            sheet, ref = A.split_sheet(pm["source"])
            if (src.sheet or ws.title, src.ref) != (sheet, ref):
                src.sheet, src.ref, changed = sheet, ref, True
        if changed and p.cache is not None:
            p.cache.refreshOnLoad = True


def _create_pivot(wb, ws, pm):
    """A pivot table created in the app (pivots.build): a pivot definition and cache in Excel's form.
    The cache holds the field items but no records (saveData=0) and asks to be rebuilt when the
    file opens (refreshOnLoad=1), so Excel always shows it computed from the source; the layout
    records (rowItems/colItems) match the cells written here, for refreshing in the app."""
    from openpyxl.pivot.cache import CacheDefinition, CacheField, CacheSource, SharedItems, WorksheetSource
    from openpyxl.pivot.fields import Index, Number, Text
    from openpyxl.pivot.table import (DataField, FieldItem, Location, PageField, PivotField, PivotTableStyle,
                                      RowColField, RowColItem, TableDefinition)
    lay = pm["layout"]
    fields = lay["fields"]
    items = lay["items"]
    sheet, ref = A.split_sheet(pm["source"])
    rows, cols, filters = pm["rows"], pm["cols"], pm.get("filters") or []
    axis = {**{f: "axisRow" for f in rows}, **{f: "axisCol" for f in cols}, **{f: "axisPage" for f in filters}}
    data_fields = {v["field"] for v in pm["values"]}

    def shared(vals):
        el = [Number(v=v) if isinstance(v, (int, float)) and not isinstance(v, bool) else Text(v=str(v)) for v in vals]
        nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
        return SharedItems(_fields=el, count=len(el), containsString=False if nums and len(nums) == len(vals) else None,
                           containsSemiMixedTypes=False if nums and len(nums) == len(vals) else None,
                           containsNumber=True if nums else None)

    cache = CacheDefinition(
        refreshOnLoad=True, saveData=False, createdVersion=6, refreshedVersion=6, minRefreshableVersion=3,
        cacheSource=CacheSource(type="worksheet", worksheetSource=WorksheetSource(ref=ref, sheet=sheet)),
        cacheFields=[CacheField(name=str(f), numFmtId=0, sharedItems=shared(items[f]) if f in axis else SharedItems())
                     for f in fields])
    pfields = []
    for f in fields:
        if f in axis:
            pfields.append(PivotField(axis=axis[f], showAll=False, compact=False, outline=False,
                                      items=[FieldItem(x=k) for k in range(len(items[f]))] + [FieldItem(t="default")]))
        else:
            pfields.append(PivotField(dataField=True if f in data_fields else None, showAll=False, compact=False, outline=False))
    fidx = {f: k for k, f in enumerate(fields)}

    def rc_items(lines, axis_fields):
        out = []
        for it in lines:
            if it["t"] == "grand":
                out.append(RowColItem(t="grand", x=[Index()]))
                continue
            if not axis_fields:  # one data field and no column field: a single empty item, <i/> (as Excel writes it)
                out.append(RowColItem())
                continue
            xs = []
            for j, v in enumerate(it["x"]):
                f = axis_fields[it["r"] + j]
                xs.append(Index(v=it.get("d", 0)) if f == "Σ" else Index(v=next(k for k, y in enumerate(items[f]) if y == v)))
            out.append(RowColItem(t=it["t"], r=it["r"] or 0, i=it.get("i") or 0, x=xs))
        return out

    col_axis = lay["colFields"]
    pt = TableDefinition(
        name=pm["name"], cacheId=1 + max([p.cacheId or 0 for w in wb.worksheets for p in getattr(w, "_pivots", [])] or [0]),
        dataCaption=PV.labels(pm.get("lang"))["values"], grandTotalCaption=PV.labels(pm.get("lang"))["grand"],
        updatedVersion=6, minRefreshableVersion=3, createdVersion=6, useAutoFormatting=True, itemPrintTitles=True,
        indent=0, compact=False, compactData=False, outline=False, outlineData=False, gridDropZones=False,
        applyNumberFormats=False, applyBorderFormats=False, applyFontFormats=False, applyPatternFormats=False,
        applyAlignmentFormats=False, applyWidthHeightFormats=True,
        location=Location(ref=pm["ref"], firstHeaderRow=lay["firstHeaderRow"], firstDataRow=lay["firstDataRow"],
                          firstDataCol=lay["firstDataCol"], rowPageCount=len(filters) or None, colPageCount=1 if filters else None),
        pivotFields=pfields,
        rowFields=[RowColField(x=fidx[f]) for f in rows],
        rowItems=rc_items(lay["rowItems"], rows),
        colFields=[RowColField(x=-2 if f == "Σ" else fidx[f]) for f in col_axis],
        colItems=rc_items(lay["colItems"], col_axis),
        pageFields=[PageField(fld=fidx[f], hier=-1) for f in filters],
        dataFields=[DataField(name=v["name"], fld=fidx[v["field"]], subtotal=v["fn"], baseField=0, baseItem=0) for v in pm["values"]],
        pivotTableStyleInfo=PivotTableStyle(name="PivotStyleLight16", showRowHeaders=True, showColHeaders=True,
                                            showRowStripes=False, showColStripes=False, showLastColumn=True),
    )
    pt.cache = cache
    ws._pivots.append(pt)
    return len(ws._pivots) - 1


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
    unsupported = []
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
            raw = None
            if isinstance(v, ArrayFormula):
                raw = v.text if v.text.startswith("=") else "=" + v.text
                item["a"] = v.ref
            elif cell.data_type == "f" and isinstance(v, str):
                raw = v
            if raw is not None:
                item["f"] = strip_prefixes(raw)
                bad = unsupported_functions(item["f"])
                if bad:
                    unsupported.append((ws.title, f"{get_column_letter(c)}{r}", bad))
                if bad or add_prefixes(item["f"], prefix_map(raw)) != raw:
                    # keep the file's exact text; written back unchanged while the formula is not edited
                    item["fo"] = raw
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
                charts.append(_chart_model(ch, i, geom, warn, theme))
            except Exception as exc:  # pragma: no cover
                warn.append(f"Could not read a chart: {exc}")
        pivots = []
        for i, p in enumerate(getattr(ws, "_pivots", None) or []):
            try:
                pivots.append(_pivot_model(p, i, ws.title))
            except Exception as exc:  # pragma: no cover
                warn.append(f"Could not read a pivot table: {exc}")
        if pivots:
            warn.append(f"Pivot tables on sheet '{ws.title}' ({', '.join(p['name'] for p in pivots)}) are kept and their areas cannot be edited. "
                        "When the source data changes, Excel refreshes them when it opens the file.")
        if getattr(ws, "_images", None):
            warn.append(f"Images on sheet '{ws.title}' are not shown but are kept when saving.")
        if ws.conditional_formatting:
            warn.append(f"Conditional formatting on sheet '{ws.title}' is not shown but is kept when saving.")

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
            "pivots": pivots,
            "showGrid": ws.sheet_view.showGridLines is not False,
        })

    for sm in sheets:  # fingerprints of the pivots' source data, compared on save
        for pm in sm["pivots"]:
            pm["srcSig"] = pivot_source_sig(sheets, pm["source"])

    names = {}
    try:
        for name, dn in wb.defined_names.items():
            if dn.attr_text:
                names[name] = strip_prefixes(dn.attr_text)
    except Exception:
        pass
    if unsupported:
        warn.insert(0, unsupported_warning(unsupported))
    if wb.chartsheets:
        warn.append("Chart sheets are not shown but are kept when saving.")
    active = wb.worksheets.index(wb.active) if wb.active in wb.worksheets else 0
    return {"sheets": sheets, "styles": styles, "activeSheet": active,
            "definedNames": names, "warnings": warn, "themeColors": theme,
            "defaultFont": {"name": default_font[0], "size": default_font[1]}}


def unsupported_warning(found) -> str:
    """found: [(sheet, 'B3', ('LAMBDA', ...)), ...] -> one user-facing warning."""
    names = []
    for _s, _ref, bad in found:
        for n in bad:
            if n not in names:
                names.append(n)
    cells = ", ".join(f"'{s}'!{ref}" for s, ref, _ in found[:5]) + (f" and {len(found) - 5} more" if len(found) > 5 else "")
    return (f"Some functions cannot be computed by this app: {', '.join(names)} ({len(found)} cells: {cells}). "
            "Those cells show the results stored in the file, and unless their formulas are edited they are saved with the original formulas. "
            "Formulas that refer to them may give different results when Excel recalculates.")


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


def _no_auto_title(chart):
    """Without <c:title>, Excel titles a one-series chart with the series name unless
    <c:autoTitleDeleted val="1"/> says not to; the preview shows no title. openpyxl's chart object
    has no setting for it, so add the element to the XML it produces."""
    from openpyxl.xml.functions import Element
    write = chart._write

    def _write():
        tree = write()
        # openpyxl builds these elements without a namespace (the chartSpace carries xmlns)
        for ns in ("", "{http://schemas.openxmlformats.org/drawingml/2006/chart}"):
            c = tree.find(ns + "chart")
            if c is not None:
                if c.find(ns + "autoTitleDeleted") is None:
                    c.insert(0, Element(ns + "autoTitleDeleted", val="1"))
                break
        return tree
    chart._write = _write


def modern_theme_xml() -> str:
    """openpyxl writes the Office 2007 theme (accent1 4F81BD, headings Cambria) into workbooks it
    creates. New workbooks get the Office 2013-2022 theme instead, the one current Excel versions
    use and the preview's colors are based on (accent1-6 4472C4 ED7D31 A5A5A5 FFC000 5B9BD5 70AD47,
    dk2 44546A, lt2 E7E6E6, hyperlink 0563C1/954F72, headings Calibri Light, body Calibri)."""
    import re
    from openpyxl.writer.theme import theme_xml
    colors_2013 = {"dk2": "44546A", "lt2": "E7E6E6", "accent1": "4472C4", "accent2": "ED7D31",
                   "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
                   "hlink": "0563C1", "folHlink": "954F72"}
    s = theme_xml
    for name, hex6 in colors_2013.items():
        s, n = re.subn(rf"(<a:{name}>\s*<a:srgbClr val=\")[0-9A-Fa-f]{{6}}", rf"\g<1>{hex6}", s, count=1)
        assert n == 1, name
    s = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*', r"\g<1>Calibri Light", s, count=1)
    return s


def _build_chart(m: dict, theme: list[str] | None = None):
    """A chart monoonom created or edited, written with the shared style (chart_style) so Excel
    draws it as the preview does. Series keep an explicit color they had in the original file."""
    from openpyxl.chart.marker import DataPoint
    from . import chart_style as CS

    kinds = {"bar": BarChart, "line": LineChart, "pie": PieChart, "doughnut": DoughnutChart,
             "area": AreaChart, "scatter": ScatterChart, "radar": RadarChart}
    main_t = m.get("type", "bar")
    groups: dict[str, list] = {}
    for s in m.get("series", []):
        groups.setdefault(s.get("type") or main_t, []).append(s)
    if main_t not in groups:
        groups = {main_t: [], **groups}
    pal = CS.palette(theme, 36)
    counter = [0]
    # palette slot = position in the chart's series list, as the preview counts (not the order the
    # series are grouped by type, and empty series still take a slot)
    slot = {id(s): i for i, s in enumerate(m.get("series", []))}

    def color_of(s):
        c = (s.get("color") or "").lstrip("#")
        return c.upper() if len(c) == 6 else pal[slot.get(id(s), 0) % len(pal)]

    def n_points(ref):
        try:
            from openpyxl.utils.cell import range_boundaries
            c1, r1, c2, r2 = range_boundaries(ref.split("!")[-1].replace("$", ""))
            return max(c2 - c1, r2 - r1) + 1
        except Exception:
            return 0

    def make(t, series):
        ch = kinds.get(t, BarChart)()
        if t == "bar":
            ch.type = m.get("barDir", "col")
            g = m.get("grouping", "clustered")
            ch.grouping = g if g in ("clustered", "stacked", "percentStacked") else "clustered"
            if ch.grouping == "clustered":
                ch.gapWidth, ch.overlap = CS.BAR_GAP, CS.BAR_OVERLAP
            else:
                ch.gapWidth, ch.overlap = CS.STACKED_GAP, CS.STACKED_OVERLAP
        elif t in ("line", "area"):
            g = m.get("grouping", "standard")
            ch.grouping = g if g in ("standard", "stacked", "percentStacked") else "standard"
        # varyColors is always written: Excel treats a missing element as "vary colors by point", so a
        # one-series column chart got a color per bar and a legend of category names. Pies vary unless
        # the file said otherwise; other charts only when the original chart did.
        if t == "scatter":
            ch.scatterStyle = "lineMarker"  # what Excel writes for "Scatter"; lines are noFill -> markers only
        pie_like = t in ("pie", "doughnut")
        vary = m.get("varyColors") is not False if pie_like else m.get("varyColors") is True
        ch.varyColors = vary
        for s in series:
            if not s.get("val") and not s.get("valLit"):
                continue
            i = counter[0]
            color = color_of(s)
            if s.get("val"):
                ser = Series(s["val"], xvalues=s.get("cat") or None) if t == "scatter" else Series(s["val"])
            else:  # values stored in the chart itself (numLit), kept as they were
                from openpyxl.chart.series import Series as _Ser
                ser = _Ser()
                setattr(ser, "yVal" if t == "scatter" else "val", NumDataSource(numLit=_num_data(s["valLit"])))
                if t == "scatter" and s.get("cat"):
                    ser.xVal = AxDataSource(numRef=NumRef(f=s["cat"]))
            if t == "scatter":
                ser.marker = CS.marker_circle(color)
            elif s.get("cat"):
                ser.cat = AxDataSource(numRef=NumRef(f=s["cat"]))
            if not s.get("cat") and s.get("catLit"):
                setattr(ser, "xVal" if t == "scatter" else "cat", _ax_lit(s["catLit"]))
            ser.graphicalProperties = CS.series_props(t, color)
            if t in ("line", "radar"):
                ser.marker = CS.marker_none()
                ser.smooth = False
            if t in ("pie", "doughnut") and not vary:
                ser.graphicalProperties = GraphicalProperties(solidFill=color, ln=CS.pie_border())
            elif t in ("pie", "doughnut"):
                # one color per slice (the file's where it had them), white borders; as many slices
                # as the value range has
                pts = s.get("pointColors") or []
                n = max(len(pts), n_points(s["val"]))
                ser.graphicalProperties = GraphicalProperties(ln=CS.pie_border())
                ser.dPt = [DataPoint(idx=k, spPr=GraphicalProperties(
                    solidFill=(pts[k].lstrip("#").upper() if k < len(pts) and pts[k] else pal[k % len(pal)]),
                    ln=CS.pie_border())) for k in range(n)]
            ser.idx = ser.order = i
            counter[0] += 1
            if s.get("nameRef"):
                ser.tx = SeriesLabel(strRef=StrRef(f=s["nameRef"]))
            elif s.get("name"):
                ser.tx = SeriesLabel(v=str(s["name"]))
            ch.series.append(ser)
        if t == "radar":
            CS.style_radar_axes(ch)
        elif t not in ("pie", "doughnut"):
            CS.style_axes(ch)
        return ch

    items = list(groups.items())
    chart = make(*items[0])
    for t, series in items[1:]:
        chart += make(t, series)
    if m.get("title"):
        chart.title = CS.title(m["title"])
    else:
        _no_auto_title(chart)
    if m.get("legend"):
        chart.legend = Legend()
        chart.legend.position = m["legend"]
        chart.legend.txPr = CS.rich()
        chart.legend.overlay = False
    else:
        chart.legend = None
    try:
        if m.get("xTitle"):
            chart.x_axis.title = CS.axis_title(m["xTitle"])
        if m.get("yTitle"):
            chart.y_axis.title = CS.axis_title(m["yTitle"])
    except Exception:
        pass
    chart.roundedCorners = False
    chart.graphical_properties = CS.chart_area()
    chart.style = None
    return chart


def _num_data(xs):
    from openpyxl.chart.data_source import NumData, NumVal
    return NumData(ptCount=len(xs), pt=[NumVal(idx=i, v=x) for i, x in enumerate(xs) if isinstance(x, (int, float))])


def _ax_lit(xs):
    from openpyxl.chart.data_source import StrData, StrVal
    if all(isinstance(x, (int, float)) or x is None for x in xs):
        return AxDataSource(numLit=_num_data(xs))
    return AxDataSource(strLit=StrData(ptCount=len(xs), pt=[StrVal(idx=i, v=str(x)) for i, x in enumerate(xs) if x is not None]))


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
        wb.loaded_theme = modern_theme_xml().encode("utf-8")  # instead of openpyxl's Office 2007 theme
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
    chart_theme = colors.parse_theme_colors(getattr(wb, "loaded_theme", None))

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

    # prefixes the opened file used (also for functions this app does not know), kept for edited formulas
    file_prefixes = {}
    for _ws, sm in ordered:
        for _r, _c, item in sm.get("cells", []):
            if item.get("fo"):
                file_prefixes.update(prefix_map(item["fo"]))

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
                f = f if f.startswith("=") else "=" + f
                fo = item.get("fo")
                text = fo if fo and strip_prefixes(fo) == f else add_prefixes(f, file_prefixes)
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
                cell.comment = Comment(item["cm"], "monoonom")

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
                ch = _build_chart(cm, chart_theme)
            ch.anchor = _anchor(cm.get("anchor", {}))
            new_charts.append(ch)
        ws._charts = new_charts
        _write_pivots(ws, sm, model["sheets"])

    active = model.get("activeSheet", 0)
    if wb.worksheets:
        wb.active = max(0, min(active, len(wb.worksheets) - 1))
        for ws in wb.worksheets:
            ws.sheet_view.tabSelected = ws is wb.active
    wb.calculation.fullCalcOnLoad = True  # Excel still recalculates on open

    _number_pivot_records(wb)
    out = io.BytesIO()
    wb.save(out)
    cached = {}
    for ws, sm in ordered:
        vals = {}
        for r, c, item in sm.get("cells", []):
            if item.get("f") and item.get("cv") is not None:
                if item["cv"] == "#NAME?" and unsupported_functions(item["f"] if item["f"].startswith("=") else "=" + item["f"]):
                    continue  # we could not calculate it; leave the result for Excel instead of storing #NAME?
                vals[f"{get_column_letter(c + 1)}{r + 1}"] = item["cv"]
        if vals:
            cached[ws.title] = vals
    data = fix_pivot_tables(fix_drawings(out.getvalue()))
    return inject_cached_values(data, cached) if cached else data


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# only a chart frame's own transform: in CT_GraphicalObjectFrame it comes right after nvGraphicFramePr
# (shapes' <a:xfrm/> in spPr/grpSpPr are left alone)
_EMPTY_XFRM = re.compile(r"(</(?:\w+:)?nvGraphicFramePr>)\s*<((?:\w+:)?)xfrm\s*/>")


def _number_pivot_records(wb):
    """openpyxl numbers pivot caches 1, 2, ... in sheet order while saving (ExcelWriter), but
    CacheDefinition._write_rels writes the link to the cache's records part before it renumbers the
    records, so once sheets or pivots change order the link names the records part the cache had
    when it was read (e.g. pivotCacheRecords1.xml) while the part is written under the new number:
    the next open fails. Give each cache's records the number its cache will get."""
    seen = []
    for ws in wb.worksheets:
        for p in getattr(ws, "_pivots", None) or []:
            if p.cache is not None and not any(p.cache is c for c in seen):
                seen.append(p.cache)
                if getattr(p.cache, "records", None) is not None:
                    p.cache.records._id = len(seen)


_PIVOT_RID = re.compile(r'(<pivotTableDefinition\b[^>]*?)\s+r:id="[^"]*"')


def fix_pivot_tables(data: bytes) -> bytes:
    """Drop the r:id attribute openpyxl puts on <pivotTableDefinition> (TableDefinition._write_rels
    stores the cache relationship id there). ECMA-376 CT_pivotTableDefinition has no such attribute,
    so the part fails the schema with it and passes without it (checked with sml.xsd); the cache is
    found through the part's relationships and the workbook's cacheId. Other parts are unchanged."""
    import zipfile

    zin = zipfile.ZipFile(io.BytesIO(data))
    names = [n for n in zin.namelist() if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml")]
    if not names:
        return data
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            raw = zin.read(info.filename)
            if info.filename in names:
                raw = _PIVOT_RID.sub(r"\1", raw.decode("utf-8"), count=1).encode("utf-8")
            zout.writestr(info, raw)
    return out.getvalue()


def fix_drawings(data: bytes) -> bytes:
    """Give every chart frame a full transform (precaution, not a known fix).

    openpyxl writes each chart's <graphicFrame> with an empty <xfrm/>. That is valid: in ECMA-376
    CT_Transform2D, <a:off> and <a:ext> are both optional (minOccurs=0), and every chart openpyxl
    writes has it. It was suspected for Excel's "Removed Part: /xl/drawings/drawingN.xml", but XSD
    checks found the scatter chart's missing <scatterStyle> instead (see docs/dev/skill-implementation.md);
    Excel reports chart-part errors as a removed drawing. This writes the form Excel itself saves,
    <xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm> (also what XlsxWriter, which is tested
    against Excel-made files, writes: xlsxwriter/drawing.py _write_xfrm); the anchor holds the position.
    Only drawing parts change; everything else stays byte-identical.
    """
    import zipfile

    zin = zipfile.ZipFile(io.BytesIO(data))
    names = [n for n in zin.namelist() if n.startswith("xl/drawings/drawing") and n.endswith(".xml")]
    if not names:
        return data
    full = lambda m: (f'{m.group(1)}<{m.group(2)}xfrm><a:off xmlns:a="{_A_NS}" x="0" y="0"/>'
                      f'<a:ext xmlns:a="{_A_NS}" cx="0" cy="0"/></{m.group(2)}xfrm>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            raw = zin.read(info.filename)
            if info.filename in names:
                raw = _EMPTY_XFRM.sub(full, raw.decode("utf-8")).encode("utf-8")
            zout.writestr(info, raw)
    return out.getvalue()


_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"}


def _xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inject_cached_values(data: bytes, cached: dict) -> bytes:
    """Write formula results (computed by the browser/Node engine) as cached <v> values.

    openpyxl writes every formula with an empty <v/>, so other tools (pandas/openpyxl data_only,
    file previews, Excel's Protected View, which does not recalculate) would show blanks.
    Only the empty <v></v> right after a formula is filled; everything else is left byte-identical.
    """
    import re
    import zipfile

    zin = zipfile.ZipFile(io.BytesIO(data))
    wbxml = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    targets.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    parts = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wbxml):
        tag = m.group(0)
        name = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]+)"', tag) or re.search(r'\bid="([^"]+)"', tag)
        if not name or not rid or rid.group(1) not in targets:
            continue
        t = targets[rid.group(1)].lstrip("/")
        part = t if t.startswith("xl/") else "xl/" + t
        sheet_name = name.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")
        parts[part] = cached.get(sheet_name)
    cell_re = re.compile(r'<c r="([A-Z]+[0-9]+)"((?:\s[^>]*?)?)><f([^>]*)>(.*?)</f><v\s*/?>(?:</v>)?</c>', re.S)

    def fill(m, vals):
        ref, attrs, fattrs, ftext = m.groups()
        if ref not in vals:
            return m.group(0)
        v = vals[ref]
        attrs = re.sub(r'\st="[^"]*"', "", attrs)
        if isinstance(v, bool):
            t, text = "b", "1" if v else "0"
        elif isinstance(v, (int, float)):
            if v != v or v in (float("inf"), float("-inf")):
                return m.group(0)
            t, text = None, repr(float(v)) if isinstance(v, float) and not float(v).is_integer() else str(int(v))
        elif isinstance(v, str) and v in _ERRORS:
            t, text = "e", v
        elif isinstance(v, str):
            t, text = "str", _xml_escape(v)
        else:
            return m.group(0)
        tattr = f' t="{t}"' if t else ""
        return f'<c r="{ref}"{attrs}{tattr}><f{fattrs}>{ftext}</f><v>{text}</v></c>'

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            raw = zin.read(info.filename)
            vals = parts.get(info.filename)
            if vals:
                raw = cell_re.sub(lambda m: fill(m, vals), raw.decode("utf-8")).encode("utf-8")
            zout.writestr(info, raw)
    return out.getvalue()
