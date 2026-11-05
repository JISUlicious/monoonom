"""A1 reference helpers (0-based rows/cols), mirroring assets/web/js/address.js."""
from __future__ import annotations

import re

MAX_ROWS, MAX_COLS = 1048576, 16384


def col_name(c: int) -> str:
    s, c = "", c + 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(name: str) -> int:
    n = 0
    for ch in name.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def a1(r, c) -> str:
    return f"{col_name(c)}{r + 1}"


def range_a1(r1, c1, r2, c2) -> str:
    return a1(r1, c1) if (r1, c1) == (r2, c2) else f"{a1(r1, c1)}:{a1(r2, c2)}"


_SHEET = re.compile(r"^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$")


def split_sheet(ref: str):
    ref = ref.strip().lstrip("=")
    m = _SHEET.match(ref)
    if m:
        return (m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2)), m.group(3)
    return None, ref


def parse_range(text: str):
    """'A1', 'A1:C3', 'A:C', '2:5' -> (r1, c1, r2, c2) normalized."""
    t = text.replace("$", "").strip().upper()
    p, _, q = t.partition(":")
    q = q or p
    if re.fullmatch(r"[A-Z]{1,3}", p) and re.fullmatch(r"[A-Z]{1,3}", q):
        c1, c2 = sorted((col_index(p), col_index(q)))
        return 0, c1, MAX_ROWS - 1, c2
    if re.fullmatch(r"\d+", p) and re.fullmatch(r"\d+", q):
        r1, r2 = sorted((int(p) - 1, int(q) - 1))
        return r1, 0, r2, MAX_COLS - 1
    mp, mq = re.fullmatch(r"([A-Z]{1,3})(\d+)", p), re.fullmatch(r"([A-Z]{1,3})(\d+)", q)
    if not mp or not mq:
        raise ValueError(f"잘못된 범위: {text}")
    r1, r2 = sorted((int(mp.group(2)) - 1, int(mq.group(2)) - 1))
    c1, c2 = sorted((col_index(mp.group(1)), col_index(mq.group(1))))
    if r1 < 0 or c2 >= MAX_COLS or r2 >= MAX_ROWS:
        raise ValueError(f"범위가 시트를 벗어납니다: {text}")
    return r1, c1, r2, c2


def quote_sheet(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_À-￿][\w.À-￿]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name) \
            and name.upper() not in ("TRUE", "FALSE"):
        return name
    return "'" + name.replace("'", "''") + "'"


_LIT = re.compile(r'("(?:[^"]|"")*")')


def rename_sheet_in_formula(formula: str, old: str, new: str) -> str:
    """Rewrite sheet prefixes Old!/'Old'! outside string literals (case-insensitive)."""
    parts = _LIT.split(formula)
    qold = re.escape(old.replace("'", "''"))
    pat = re.compile(rf"(?:'{qold}'|(?<![\w.À-￿']){re.escape(old)})!", re.IGNORECASE)
    rep = quote_sheet(new) + "!"
    return "".join(p if i % 2 else pat.sub(lambda m: rep, p) for i, p in enumerate(parts))
