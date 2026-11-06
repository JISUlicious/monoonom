"""Fonts installed on this computer, for the font picker (GET /api/fonts) and `xl.py fonts`.

The server runs on the user's machine, so it can list the fonts without the browser's Local Font
Access permission prompt (which iframe previews may not allow at all). Each font file's OpenType
`name` table gives the family name, in English (e.g. "Malgun Gothic") and in Korean where the font has one (as Korean Excel shows it);
both work as a cell's font name and in CSS.
Only the table directory and the name table are read, not whole files (CJK fonts are 10-20 MB).
Format: OpenType spec, "name — Naming Table" and "Font Collections" (TTC header).
"""
from __future__ import annotations

import os
import struct
import sys
import time

KO, EN = 0x0412, 0x0409
_cache: dict | None = None


def font_dirs() -> list[str]:
    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        win = os.environ.get("WINDIR", r"C:\Windows")
        local = os.environ.get("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
        return [os.path.join(win, "Fonts"), os.path.join(local, "Microsoft", "Windows", "Fonts")]
    if sys.platform == "darwin":
        return ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/Library/Fonts", os.path.join(home, "Library", "Fonts")]
    return ["/usr/share/fonts", "/usr/local/share/fonts", os.path.join(home, ".fonts"), os.path.join(home, ".local", "share", "fonts")]


def _names(fh, base: int) -> dict:
    """{languageID: family} from one font's name table (typographic family 16 over family 1)."""
    fh.seek(base + 4)
    (num_tables,) = struct.unpack(">H", fh.read(2))
    fh.seek(base + 12)
    name_off = name_len = None
    for _ in range(min(num_tables, 200)):
        tag, _chk, off, length = struct.unpack(">4sIII", fh.read(16))
        if tag == b"name":
            name_off, name_len = off, length
            break
    if name_off is None:
        return {}
    fh.seek(name_off)
    data = fh.read(min(name_len, 1 << 20))
    _fmt, count, str_off = struct.unpack(">HHH", data[:6])
    found: dict = {}  # (nameID, lang) -> text
    for i in range(min(count, 2000)):
        rec = data[6 + 12 * i: 18 + 12 * i]
        if len(rec) < 12:
            break
        plat, enc, lang, nid, length, off = struct.unpack(">HHHHHH", rec)
        if nid not in (1, 16):
            continue
        raw = data[str_off + off: str_off + off + length]
        if plat == 3:  # Windows, UTF-16BE
            try:
                text = raw.decode("utf-16-be")
            except UnicodeDecodeError:
                continue
        elif plat == 1 and lang == 0:  # Macintosh, Roman, English
            text, lang = raw.decode("mac_roman", "replace"), EN
        elif plat == 0:  # Unicode
            try:
                text, lang = raw.decode("utf-16-be"), EN
            except UnicodeDecodeError:
                continue
        else:
            continue
        text = text.strip("\x00 ").strip()
        if text:
            found[(nid, lang)] = text
    out = {}
    for lang in {lang for _, lang in found}:
        out[lang] = found.get((16, lang)) or found.get((1, lang))
    return out


def families_in(path: str) -> list[dict]:
    """[{"ko": name or None, "en": name}] for each font in a .ttf/.otf/.ttc/.otc file."""
    out = []
    try:
        with open(path, "rb") as fh:
            head = fh.read(12)
            if head[:4] == b"ttcf":
                (n,) = struct.unpack(">I", head[8:12])
                offs = struct.unpack(f">{min(n, 64)}I", fh.read(4 * min(n, 64)))
            else:
                offs = (0,)
            for base in offs:
                nm = _names(fh, base)
                en = nm.get(EN) or next(iter(nm.values()), None)
                if en:
                    out.append({"ko": nm.get(KO), "en": en})
    except (OSError, struct.error, ValueError):
        pass
    return out


def list_fonts(refresh: bool = False, lang: str = "en") -> dict:
    """{"fonts": [{"name": shown name, "names": [all names]}], "files", "source", "ms"}: deduplicated and
    sorted. The shown name is the English family name; with lang "ko" it is the Korean one where the
    font has it (as Korean Excel shows it), and fonts with Korean names come first."""
    global _cache
    if _cache is None or refresh:
        t0 = time.time()
        fams: dict = {}
        files = 0
        for d in font_dirs():
            for root, _dirs, names in os.walk(d):
                for n in names:
                    if not n.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
                        continue
                    files += 1
                    if files > 6000:
                        break
                    for f in families_in(os.path.join(root, n)):
                        if f["en"].startswith("."):  # macOS system UI fonts (".SF NS"): not for documents
                            continue
                        e = fams.setdefault(f["en"].lower(), {"en": f["en"], "ko": None})
                        e["ko"] = e["ko"] or f["ko"]
        _cache = {"families": list(fams.values()), "files": files, "ms": round((time.time() - t0) * 1000)}
    ko = lang == "ko"
    fonts = [{"name": (f["ko"] if ko and f["ko"] else f["en"]), "names": [x for x in (f["ko"], f["en"]) if x]}
             for f in _cache["families"]]
    fonts.sort(key=lambda f: (ko and not any("\uac00" <= ch <= "\ud7a3" for ch in f["name"]), f["name"].lower()))
    return {"fonts": fonts, "files": _cache["files"], "source": "system", "ms": _cache["ms"]}
