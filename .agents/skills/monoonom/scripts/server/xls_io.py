"""Legacy .xls (BIFF8) support: read only, with xlrd.

.xls files are never written: saving an .xls document writes a new .xlsx next to it (same name)
and leaves the original untouched. xlrd does not expose cell formulas, only their values, so
writing back to .xls would also have replaced the file's formulas with values.
"""
from __future__ import annotations

import xlrd

from .formula_compat import strip_prefixes
from .units import DEFAULT_COL_PX, DEFAULT_ROW_PX, colw_to_px, pt_to_px

BORDER_STYLES = ["", "thin", "medium", "dashed", "dotted", "thick", "double", "hair",
                 "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
                 "mediumDashDotDot", "slantDashDot"]
HALIGN = {1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "centerContinuous"}
VALIGN = {0: "top", 1: "center", 3: "justify"}


def _hex(rgb):
    return "#%02X%02X%02X" % rgb if rgb else None


def _style(book, xf_index, cache):
    if xf_index in cache:
        return cache[xf_index]
    xf = book.xf_list[xf_index]
    cmap = book.colour_map
    st = {}
    font = book.font_list[xf.font_index]
    if font.bold:
        st["b"] = True
    if font.italic:
        st["i"] = True
    if font.underline_type:
        st["u"] = "double" if font.underline_type in (2, 34) else True
    if font.struck_out:
        st["st"] = True
    f0 = book.font_list[0]
    if font.height and font.height != f0.height:
        st["fs"] = font.height / 20
    if font.name and font.name != f0.name:
        st["ff"] = font.name
    fc = _hex(cmap.get(font.colour_index))
    if fc and fc != "#000000":
        st["fc"] = fc
    if xf.background.fill_pattern:
        bg = _hex(cmap.get(xf.background.pattern_colour_index))
        if bg:
            st["bg"] = bg
    al = xf.alignment
    if al.hor_align in HALIGN:
        st["ha"] = HALIGN[al.hor_align]
    if al.vert_align in VALIGN:
        st["va"] = VALIGN[al.vert_align]
    if al.text_wrapped:
        st["wrap"] = True
    if al.indent_level:
        st["ind"] = al.indent_level
    b = xf.border
    for key, side in (("bl", "left"), ("br", "right"), ("bt", "top"), ("bb", "bottom")):
        ls = getattr(b, f"{side}_line_style")
        if ls:
            e = {"s": BORDER_STYLES[ls] if ls < len(BORDER_STYLES) else "thin"}
            c = _hex(cmap.get(getattr(b, f"{side}_colour_index")))
            if c:
                e["c"] = c
            st[key] = e
    fmt = book.format_map.get(xf.format_key)
    if fmt and fmt.format_str and fmt.format_str != "General":
        st["nf"] = fmt.format_str
    cache[xf_index] = st
    return st


def read_xls(data: bytes) -> dict:
    book = xlrd.open_workbook(file_contents=data, formatting_info=True)
    styles, style_ids, xf_cache = [], {}, {}
    sheets = []
    date_shift = 1462 if book.datemode == 1 else 0

    def style_id(xf_index):
        st = _style(book, xf_index, xf_cache)
        if not st:
            return None
        key = repr(sorted(st.items()))
        if key not in style_ids:
            style_ids[key] = len(styles)
            styles.append(st)
        return style_ids[key]

    for sh in book.sheets():
        dcol = colw_to_px(sh.defcolwidth + 0.43) if sh.defcolwidth else (
            colw_to_px(sh.standardwidth / 256) if sh.standardwidth else DEFAULT_COL_PX)
        drow = pt_to_px(sh.default_row_height / 20) if sh.default_row_height else DEFAULT_ROW_PX
        cols = {}
        for c, info in sh.colinfo_map.items():
            e = {"w": colw_to_px(info.width / 256)}
            if info.hidden:
                e["hidden"] = True
            cols[str(c)] = e
        rows = {}
        for r, info in sh.rowinfo_map.items():
            e = {}
            if info.height_mismatch or info.height != sh.default_row_height:
                e["h"] = pt_to_px(info.height / 20)
            if info.hidden:
                e["hidden"] = True
            if e:
                rows[str(r)] = e
        cells = []
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                ct = sh.cell_type(r, c)
                item = {}
                if ct == xlrd.XL_CELL_TEXT:
                    item["v"] = sh.cell_value(r, c)
                elif ct == xlrd.XL_CELL_NUMBER:
                    item["v"] = sh.cell_value(r, c)
                elif ct == xlrd.XL_CELL_DATE:
                    item["v"] = sh.cell_value(r, c) + date_shift
                elif ct == xlrd.XL_CELL_BOOLEAN:
                    item["v"] = bool(sh.cell_value(r, c))
                elif ct == xlrd.XL_CELL_ERROR:
                    item["v"] = xlrd.error_text_from_code.get(sh.cell_value(r, c), "#N/A")
                    item["t"] = "e"
                if ct != xlrd.XL_CELL_EMPTY:
                    sid = style_id(sh.cell_xf_index(r, c))
                    if sid is not None:
                        item["s"] = sid
                if isinstance(item.get("v"), float) and item["v"].is_integer() and ct == xlrd.XL_CELL_NUMBER:
                    item["v"] = int(item["v"])
                if item:
                    cells.append([r, c, item])
        merges = []
        for rlo, rhi, clo, chi in sh.merged_cells:
            merges.append(f"{_col(clo)}{rlo + 1}:{_col(chi - 1)}{rhi}")
        freeze = None
        if sh.panes_are_frozen and (sh.horz_split_pos or sh.vert_split_pos):
            freeze = f"{_col(sh.vert_split_pos or 0)}{(sh.horz_split_pos or 0) + 1}"
        sheets.append({
            "name": sh.name, "origName": sh.name,
            "state": "hidden" if sh.visibility else "visible", "tabColor": None,
            "defaultColW": dcol, "defaultRowH": drow, "cols": cols, "rows": rows,
            "cells": cells, "merges": merges, "freeze": freeze, "autoFilter": None,
            "charts": [], "showGrid": bool(sh.show_grid_lines),
        })
    warns = ["xls (Excel 97-2003) files are read-only. Formulas show their computed values (an xlrd limit), "
             "and saving writes a .xlsx with the same name. The original .xls is not changed."]
    names = {}
    for nm in book.name_obj_list:
        try:
            if nm.formula_text and not nm.builtin:
                names[nm.name] = strip_prefixes(nm.formula_text)
        except Exception:
            pass
    f0 = book.font_list[0]
    return {"sheets": sheets, "styles": styles, "activeSheet": 0,
            "definedNames": names, "warnings": warns,
            "defaultFont": {"name": f0.name, "size": f0.height / 20}}


def _col(c: int) -> str:
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s
