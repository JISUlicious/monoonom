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


def lum(hex6: str, mod: int | None, off: int | None) -> str:
    """DrawingML lumMod/lumOff (1/1000 %): L' = L * mod + off in HSL (ECMA-376 §20.1.2.3.20-21)."""
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * ((mod if mod is not None else 100000) / 100000) + (off or 0) / 100000
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def shade_tint(hex6: str, shade: int | None = None, tint: int | None = None) -> str:
    """DrawingML shade/tint (1/1000 %): a 40 % shade is 40 % of the color and 60 % black, a 40 % tint
    40 % of the color and 60 % white, blended in linear RGB (ECMA-376 §20.1.2.3.31 shade, .34 tint)."""
    def lin(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    def srgb(c):
        c = max(0.0, min(1.0, c))
        return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055
    out = []
    for i in (0, 2, 4):
        c = lin(int(hex6[i:i + 2], 16) / 255)
        if shade is not None:
            c = c * shade / 100000
        if tint is not None:
            c = c * tint / 100000 + (1 - tint / 100000)
        out.append(round(srgb(c) * 255))
    return "%02X%02X%02X" % tuple(out)


SCHEME_INDEX = {"lt1": 0, "bg1": 0, "dk1": 1, "tx1": 1, "lt2": 2, "bg2": 2, "dk2": 3, "tx2": 3,
                "accent1": 4, "accent2": 5, "accent3": 6, "accent4": 7, "accent5": 8, "accent6": 9,
                "hlink": 10, "folHlink": 11}


def drawing_color(choice, theme: list[str]) -> str | None:
    """DrawingML color (openpyxl ColorChoice: srgbClr, schemeClr with lumMod/lumOff/shade/tint,
    sysClr) -> 'RRGGBB', or None when there is none."""
    if choice is None:
        return None
    try:
        if isinstance(choice, str):
            return choice[-6:].upper() if len(choice) >= 6 else None
        if getattr(choice, "srgbClr", None):
            v = choice.srgbClr
            v = v if isinstance(v, str) else getattr(v, "val", None)
            return v[-6:].upper() if v else None
        sc = getattr(choice, "schemeClr", None)
        if sc is not None and sc.val in SCHEME_INDEX:
            v = (theme or DEFAULT_THEME)[SCHEME_INDEX[sc.val]]
            if sc.lumMod is not None or sc.lumOff is not None:
                v = lum(v, sc.lumMod, sc.lumOff)
            if sc.shade is not None:
                v = shade_tint(v, shade=sc.shade)
            if sc.tint is not None:
                v = shade_tint(v, tint=sc.tint)
            return v
        sys_c = getattr(choice, "sysClr", None)
        if sys_c is not None and getattr(sys_c, "lastClr", None):
            return sys_c.lastClr[-6:].upper()
    except Exception:  # pragma: no cover - unusual color markup
        return None
    return None


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
