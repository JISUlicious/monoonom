"""Legacy .xls (BIFF8) support: read with xlrd, write with xlwt.

Limits (format/library, see [MS-XLS]): 65,536 rows x 256 columns, 56-color palette,
no charts/autofilter written; xlrd does not expose cell formulas (only their values).
"""
from __future__ import annotations

import io

import struct

import xlrd
import xlwt
from xlwt import BIFFRecords
from xlwt.Cell import FormulaCell
from xlrd.formatting import excel_default_palette_b8

from .formula_compat import strip_prefixes
from .units import DEFAULT_COL_PX, DEFAULT_ROW_PX, colw_to_px, pt_to_px, px_to_pt

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
    warns = ["xls 형식은 수식 텍스트를 읽을 수 없어 계산된 값으로 표시됩니다(xlrd 제한). "
             "수식을 유지하려면 xlsx로 저장하세요."]
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


# --------------------------------------------------------------------------- write

def _formula_biff(self):
    """xlwt always stores an empty-string cached result; store the computed value instead
    (FORMULA record FormulaValue, [MS-XLS] 2.5.133) so readers without a calc engine see it."""
    rec = BIFFRecords.FormulaRecord(self.rowx, self.colx, self.xf_idx, self.frmla.rpn(), self.calc_flags | 2).get()
    v = getattr(self.frmla, "_cached", None)
    if isinstance(v, bool):
        res = struct.pack("<BBBBHH", 1, 0, int(v), 0, 0, 0xFFFF)
    elif isinstance(v, (int, float)):
        res = struct.pack("<d", float(v))
    else:
        return rec
    return rec[:10] + res + rec[18:]


FormulaCell.get_biff_data = _formula_biff

_PALETTE = list(enumerate(excel_default_palette_b8, start=8))


def _nearest_colour(hex_):
    if not hex_:
        return None
    h = hex_.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    best = min(_PALETTE, key=lambda p: (p[1][0] - r) ** 2 + (p[1][1] - g) ** 2 + (p[1][2] - b) ** 2)
    return best[0]


_HALIGN_W = {"left": 1, "center": 2, "right": 3, "fill": 4, "justify": 5, "centerContinuous": 6}
_VALIGN_W = {"top": 0, "center": 1, "bottom": 2, "justify": 3}


_DEFAULT_FONT = ["Arial", 10.0]


def _xf(st: dict, cache: dict):
    key = repr(sorted(st.items()))
    if key in cache:
        return cache[key]
    x = xlwt.XFStyle()
    f = xlwt.Font()
    f.name = st.get("ff", _DEFAULT_FONT[0])
    f.height = int(st.get("fs", _DEFAULT_FONT[1]) * 20)
    f.bold = bool(st.get("b"))
    f.italic = bool(st.get("i"))
    f.underline = 2 if st.get("u") == "double" else (1 if st.get("u") else 0)
    f.struck_out = bool(st.get("st"))
    if st.get("fc"):
        f.colour_index = _nearest_colour(st["fc"])
    x.font = f
    if st.get("bg"):
        p = xlwt.Pattern()
        p.pattern = xlwt.Pattern.SOLID_PATTERN
        p.pattern_fore_colour = _nearest_colour(st["bg"])
        x.pattern = p
    a = xlwt.Alignment()
    a.horz = _HALIGN_W.get(st.get("ha"), 0)
    a.vert = _VALIGN_W.get(st.get("va", "bottom"), 2)
    a.wrap = 1 if st.get("wrap") else 0
    a.inde = int(st.get("ind", 0))
    x.alignment = a
    bd = xlwt.Borders()
    for key, side in (("bl", "left"), ("br", "right"), ("bt", "top"), ("bb", "bottom")):
        e = st.get(key)
        if e:
            s = e.get("s", "thin")
            setattr(bd, side, BORDER_STYLES.index(s) if s in BORDER_STYLES else 1)
            setattr(bd, f"{side}_colour", _nearest_colour(e.get("c", "#000000")))
    x.borders = bd
    x.num_format_str = st.get("nf", "General")
    cache[key] = x
    return x


def write_xls(model: dict) -> tuple[bytes, list[str]]:
    warns = []
    dfont = model.get("defaultFont") or {}
    _DEFAULT_FONT[:] = [dfont.get("name") or "Arial", float(dfont.get("size") or 10)]
    wb = xlwt.Workbook(encoding="utf-8")
    wb.get_default_style().font.name = _DEFAULT_FONT[0]
    wb.get_default_style().font.height = int(_DEFAULT_FONT[1] * 20)
    styles = model.get("styles", [])
    cache: dict = {}
    default = xlwt.XFStyle()
    for sm in model["sheets"]:
        ws = wb.add_sheet(sm["name"][:31], cell_overwrite_ok=True)
        if sm.get("charts"):
            warns.append(f"'{sm['name']}': xls 형식에는 차트를 저장할 수 없습니다.")
        if sm.get("autoFilter"):
            warns.append(f"'{sm['name']}': xls 저장 시 자동 필터 설정은 제외되고 숨김 행만 유지됩니다.")
        for key, e in sm.get("cols", {}).items():
            c = int(key)
            if c > 255:
                continue
            col = ws.col(c)
            if "w" in e:
                col.width = int(max(0, (e["w"] - 5) / 7) * 256)
            if e.get("hidden"):
                col.hidden = True
        for key, e in sm.get("rows", {}).items():
            r = int(key)
            if r > 65535:
                continue
            row = ws.row(r)
            if "h" in e:
                row.height_mismatch = True
                row.height = int(px_to_pt(e["h"]) * 20)
            if e.get("hidden"):
                row.hidden = True
        dropped = 0
        for r, c, item in sm.get("cells", []):
            if r > 65535 or c > 255:
                dropped += 1
                continue
            si = item.get("s")
            style = _xf(styles[si], cache) if isinstance(si, int) and 0 <= si < len(styles) else default
            v = item.get("v")
            f = item.get("f")
            if f:
                try:
                    fo = xlwt.Formula(f.lstrip("="))
                    fo._cached = item.get("cv")
                    ws.write(r, c, fo, style)
                    continue
                except Exception:
                    warns.append(f"'{sm['name']}'!{_col(c)}{r + 1}: xls에서 지원하지 않는 수식이라 값으로 저장했습니다.")
                    v = item.get("cv", v)
            if v is None:
                if style is not default:
                    ws.write(r, c, None, style)  # styled blank cell
            else:
                ws.write(r, c, v, style)
        if dropped:
            warns.append(f"'{sm['name']}': xls 한계(65,536행 x 256열)를 넘는 {dropped}개 셀은 저장되지 않았습니다.")
        for ref in sm.get("merges", []):
            a, _, b = ref.partition(":")
            r1, c1 = _parse(a)
            r2, c2 = _parse(b or a)
            if r2 <= 65535 and c2 <= 255:
                ws.merge(r1, r2, c1, c2)
        fz = sm.get("freeze")
        if fz:
            r, c = _parse(fz)
            ws.set_panes_frozen(True)
            ws.set_horz_split_pos(r)
            ws.set_vert_split_pos(c)
            ws.set_remove_splits(True)
    if len(warns) > 30:
        warns = warns[:30] + [f"... 외 {len(warns) - 30}건"]
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), warns


def _parse(a1: str) -> tuple[int, int]:
    a1 = a1.replace("$", "")
    i = 0
    while i < len(a1) and a1[i].isalpha():
        i += 1
    col = 0
    for ch in a1[:i].upper():
        col = col * 26 + (ord(ch) - 64)
    return int(a1[i:]) - 1, col - 1
