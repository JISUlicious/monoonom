"""Unit conversions between SpreadsheetML and screen pixels (96 dpi).

Column width in OOXML is measured in characters of the maximum digit width (MDW) of the
default font (ECMA-376 Part 1 §18.3.1.13). For Calibri 11 the MDW is 7 px, so
px = width*7 + 5 (padding) — 8.43 chars -> 64 px, Excel's default.
Row heights are in points: px = pt * 96 / 72.
"""
MDW = 7
EMU_PER_PX = 9525  # 914400 EMU per inch / 96 dpi

DEFAULT_COL_PX = 64
DEFAULT_ROW_PX = 20


def colw_to_px(w: float) -> int:
    return max(0, round(w * MDW + 5)) if w else 0


def px_to_colw(px: float) -> float:
    return round(max(0.0, (px - 5) / MDW), 2)


def pt_to_px(pt: float) -> int:
    return round(pt * 96 / 72)


def px_to_pt(px: float) -> float:
    return round(px * 72 / 96, 2)
