"""Color resolution for SpreadsheetML (rgb / indexed / theme+tint).

Theme index order and tint math follow ECMA-376 Part 1 §18.8.3 (color) and §18.8.19 (tint):
index 0..3 map to lt1, dk1, lt2, dk2 (Excel swaps dk/lt), 4..9 accent1..6, 10 hlink, 11 folHlink.
"""
from __future__ import annotations

import colorsys
import re

from openpyxl.styles.colors import COLOR_INDEX

DEFAULT_THEME = [
    "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
    "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72",
]


def parse_theme_colors(theme_xml: bytes | None) -> list[str]:
    if not theme_xml:
        return list(DEFAULT_THEME)
    try:
        text = theme_xml.decode("utf-8", "ignore")
        m = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", text, re.S)
        if not m:
            return list(DEFAULT_THEME)
        body = m.group(1)
        found = {}
        for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                     "accent4", "accent5", "accent6", "hlink", "folHlink"):
            mm = re.search(rf"<a:{name}>(.*?)</a:{name}>", body, re.S)
            if not mm:
                continue
            c = re.search(r'(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"', mm.group(1))
            if c:
                found[name] = c.group(1).upper()
        order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink"]
        return [found.get(n, DEFAULT_THEME[i]) for i, n in enumerate(order)]
    except Exception:  # pragma: no cover - defensive
        return list(DEFAULT_THEME)


def apply_tint(hex6: str, tint: float) -> str:
    if not tint:
        return hex6
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def resolve(color, theme: list[str]) -> str | None:
    """openpyxl Color -> '#RRGGBB' or None (auto / unset)."""
    if color is None:
        return None
    try:
        t = color.type
        if t == "rgb":
            v = color.rgb
            if not isinstance(v, str) or len(v) < 6:
                return None
            v = v[-6:]
            # 'FF000000' alpha-ARGB; ignore alpha
        elif t == "indexed":
            idx = color.indexed
            if idx is None or idx >= 64:  # 64 = system foreground/auto
                return None
            v = COLOR_INDEX[idx][-6:]
        elif t == "theme":
            idx = color.theme
            if idx is None or idx >= len(theme):
                return None
            v = theme[idx]
        else:
            return None
        tint = getattr(color, "tint", 0) or 0
        return "#" + apply_tint(v.upper(), tint)
    except Exception:
        return None


def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
