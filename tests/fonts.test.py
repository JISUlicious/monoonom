"""Installed-font listing (server/fonts.py): python3 tests/fonts.test.py

Synthetic font files (just the parts the reader uses: the table directory and an OpenType `name`
table, OpenType spec "name — Naming Table"; a TTC header for collections) check that the Korean
and English family names are read, that the typographic family (name ID 16) wins over ID 1, and
that broken files are skipped. Then the real scan of this computer's font folders runs.
"""
import os
import struct
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts"))
from server import fonts as F  # noqa: E402

fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def name_table(records):
    """records: [(platform, encoding, language, nameID, text)]"""
    strings, recs = b"", []
    for plat, enc, lang, nid, text in records:
        raw = text.encode("utf-16-be") if plat in (0, 3) else text.encode("mac_roman")
        recs.append(struct.pack(">HHHHHH", plat, enc, lang, nid, len(raw), len(strings)))
        strings += raw
    head = struct.pack(">HHH", 0, len(records), 6 + 12 * len(records))
    return head + b"".join(recs) + strings


def font(records, base=0):
    """One font: offset table with a single 'name' table right after the directory."""
    table = name_table(records)
    off = base + 12 + 16
    return struct.pack(">IHHHH", 0x00010000, 1, 16, 0, 0) + struct.pack(">4sIII", b"name", 0, off, len(table)) + table


malgun = [(3, 1, 0x0409, 1, "Malgun Gothic"), (3, 1, 0x0412, 1, "맑은 고딕"), (1, 0, 0, 1, "Malgun Gothic")]
typo = [(3, 1, 0x0409, 1, "Noto Sans KR Medium"), (3, 1, 0x0409, 16, "Noto Sans KR")]
d = tempfile.mkdtemp()
open(os.path.join(d, "malgun.ttf"), "wb").write(font(malgun))
open(os.path.join(d, "noto.otf"), "wb").write(font(typo))
open(os.path.join(d, "broken.ttf"), "wb").write(b"\x00\x01\x00\x00garbage")
# a collection: TTC header + two fonts
f1 = font(malgun, base=0)  # placeholder to know the size
n1 = font(malgun, base=20)
n2 = font([(3, 1, 0x0409, 1, "Batang"), (3, 1, 0x0412, 1, "바탕")], base=20 + len(n1))
open(os.path.join(d, "pair.ttc"), "wb").write(b"ttcf" + struct.pack(">HHI", 1, 0, 2) + struct.pack(">II", 20, 20 + len(n1)) + n1 + n2)

check("Korean and English family names of one font", F.families_in(os.path.join(d, "malgun.ttf")) == [{"ko": "맑은 고딕", "en": "Malgun Gothic"}],
      str(F.families_in(os.path.join(d, "malgun.ttf"))))
check("typographic family (name ID 16) over the style-specific family (ID 1)", F.families_in(os.path.join(d, "noto.otf")) == [{"ko": None, "en": "Noto Sans KR"}],
      str(F.families_in(os.path.join(d, "noto.otf"))))
check("font collection (.ttc): every font in it", [x["ko"] for x in F.families_in(os.path.join(d, "pair.ttc"))] == ["맑은 고딕", "바탕"],
      str(F.families_in(os.path.join(d, "pair.ttc"))))
check("a broken file is skipped, not an error", F.families_in(os.path.join(d, "broken.ttf")) == [])

orig = F.font_dirs
F.font_dirs = lambda: [d]
r = F.list_fonts(refresh=True)
F.font_dirs = orig
names = [f["name"] for f in r["fonts"]]
check("list: shown in Korean where there is a Korean name, Korean names first, both names kept",
      names[:2] == ["맑은 고딕", "바탕"] and "Noto Sans KR" in names and r["fonts"][0]["names"] == ["맑은 고딕", "Malgun Gothic"], str(r["fonts"]))
real = F.list_fonts(refresh=True)
check("this computer's font folders are read", real["files"] >= 0 and all(f["name"] for f in real["fonts"]), f"{len(real['fonts'])} fonts from {real['files']} files in {real['ms']} ms")
print(f"{fails} FAILED" if fails else "ALL PASSED")
sys.exit(1 if fails else 0)
