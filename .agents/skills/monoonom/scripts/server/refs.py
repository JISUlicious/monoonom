"""Formula reference rewriting — a faithful port of assets/web/js/formula.js (tokenize + rewriteRefs).

Used by the server for agent row/column insert/delete, sort and fill so formulas end up identical to
what the browser would produce. tests/refs_parity.test.py checks both implementations against each
other on random formulas.
"""
from __future__ import annotations

import re

from .a1 import MAX_COLS, MAX_ROWS, col_index, col_name, quote_sheet

_U = "À-￿"
SHEET = rf"(?:'((?:[^']|'')+)'|([A-Za-z_{_U}][A-Za-z0-9_.{_U}]*))!"
CELL = r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)"
REF_RE = re.compile(rf"(?:{SHEET})?(?:{CELL}(?::{CELL})?|(\$?)([A-Za-z]{{1,3}}):(\$?)([A-Za-z]{{1,3}})|(\$?)(\d+):(\$?)(\d+))")
NUM_RE = re.compile(r"(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")
ROWRANGE_RE = re.compile(r"\d+:\$?\d")
IDENT_RE = re.compile(rf"[A-Za-z_\\{_U}][A-Za-z0-9_.{_U}]*")
SHEET_REF_ERR = re.compile(rf"{SHEET}#REF!")
REF_START = re.compile(rf"[A-Za-z_$'{_U}0-9]")
AFTER_BAD = re.compile(rf"[A-Za-z0-9_(.{_U}!]")
ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"]


def _ref_from_match(m):
    g = m.groups()
    sheet = g[0].replace("''", "'") if g[0] is not None else g[1]
    if g[3]:
        a = {"cAbs": bool(g[2]), "c": col_index(g[3]), "rAbs": bool(g[4]), "r": int(g[5]) - 1}
        b = {"cAbs": bool(g[6]), "c": col_index(g[7]), "rAbs": bool(g[8]), "r": int(g[9]) - 1} if g[7] else None
        if a["c"] >= MAX_COLS or a["r"] < 0 or a["r"] >= MAX_ROWS or (b and (b["c"] >= MAX_COLS or b["r"] < 0 or b["r"] >= MAX_ROWS)):
            return None
        return {"kind": "cell", "sheet": sheet, "a": a, "b": b}
    if g[11]:
        a = {"cAbs": bool(g[10]), "c": col_index(g[11]), "rAbs": False, "r": 0}
        b = {"cAbs": bool(g[12]), "c": col_index(g[13]), "rAbs": False, "r": MAX_ROWS - 1}
        if a["c"] >= MAX_COLS or b["c"] >= MAX_COLS:
            return None
        return {"kind": "col", "sheet": sheet, "a": a, "b": b}
    a = {"rAbs": bool(g[14]), "r": int(g[15]) - 1, "cAbs": False, "c": 0}
    b = {"rAbs": bool(g[16]), "r": int(g[17]) - 1, "cAbs": False, "c": MAX_COLS - 1}
    if a["r"] < 0 or b["r"] < 0:
        return None
    return {"kind": "row", "sheet": sheet, "a": a, "b": b}


def ref_tokens(formula: str):
    """[(start, end, ref)] for every reference token, using the same scan as the JS tokenizer."""
    off = 1 if formula.startswith("=") else 0
    s = formula[off:]
    out = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch.isspace():
            i += 1
            continue
        if ch == '"':
            j = i + 1
            while True:
                if j >= n:
                    return out  # unterminated string: JS throws -> formula left unchanged
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            i = j + 1
            continue
        if ch == "#":
            err = next((e for e in ERRORS if s[i:i + len(e)].upper() == e), None)
            if err:
                i += len(err)
                continue
        m = NUM_RE.match(s, i)
        if m and not ROWRANGE_RE.match(s, i):
            i = m.end()
            continue
        if REF_START.match(ch):
            m = REF_RE.match(s, i)
            if m and m.end() > i:
                after = s[m.end()] if m.end() < n else ""
                if not AFTER_BAD.match(after or " "):
                    ref = _ref_from_match(m)
                    if ref:
                        out.append((i + off, m.end() + off, ref))
                        i = m.end()
                        continue
            m = SHEET_REF_ERR.match(s, i)
            if m:
                i = m.end()
                continue
            m = IDENT_RE.match(s, i)
            if m:
                i = m.end()
                continue
        if ch not in "+-*/^&=<>%:(),;{}!":
            return []  # JS tokenizer throws on unknown characters -> formula left unchanged
        i += 1
    return out


def ref_to_string(ref) -> str:
    sh = quote_sheet(ref["sheet"]) + "!" if ref["sheet"] is not None else ""
    cp = lambda p: ("$" if p["cAbs"] else "") + col_name(p["c"])
    rp = lambda p: ("$" if p["rAbs"] else "") + str(p["r"] + 1)
    if ref["kind"] == "col":
        return f"{sh}{cp(ref['a'])}:{cp(ref['b'])}"
    if ref["kind"] == "row":
        return f"{sh}{rp(ref['a'])}:{rp(ref['b'])}"
    return sh + cp(ref["a"]) + rp(ref["a"]) + (":" + cp(ref["b"]) + rp(ref["b"]) if ref["b"] else "")


def rewrite_refs(formula: str, fn) -> str:
    toks = ref_tokens(formula)
    if not toks:
        return formula
    out, last = [], 0
    for s, e, ref in toks:
        nr = fn({**ref, "a": dict(ref["a"]), "b": dict(ref["b"]) if ref["b"] else None})
        out.append(formula[last:s])
        if nr is None:
            out.append((quote_sheet(ref["sheet"]) + "!" if ref["sheet"] is not None else "") + "#REF!")
        else:
            out.append(ref_to_string(nr))
        last = e
    out.append(formula[last:])
    return "".join(out)


def translate(formula: str, dr: int, dc: int) -> str:
    """Copy/fill semantics: shift relative parts by (dr, dc)."""
    if not dr and not dc:
        return formula

    def fn(ref):
        for p in (ref["a"], ref["b"]):
            if not p:
                continue
            if ref["kind"] != "row" and not p["cAbs"]:
                p["c"] += dc
            if ref["kind"] != "col" and not p["rAbs"]:
                p["r"] += dr
            if p["c"] < 0 or p["r"] < 0 or p["c"] >= MAX_COLS or p["r"] >= MAX_ROWS:
                return None
        return ref
    return rewrite_refs(formula, fn)


def adjust(formula: str, here: str, target: str, axis: str, at: int, count: int) -> str:
    """Insert (count>0) / delete (count<0) rows or columns at `at` on sheet `target`."""
    def fn(ref):
        sheet = ref["sheet"] if ref["sheet"] is not None else here
        if (sheet or "").lower() != (target or "").lower():
            return ref
        if (axis == "row" and ref["kind"] == "col") or (axis == "col" and ref["kind"] == "row"):
            return ref
        k = "r" if axis == "row" else "c"
        pts = [ref["a"], ref["b"]] if ref["b"] else [ref["a"]]
        if count > 0:
            for p in pts:
                if p[k] >= at:
                    p[k] += count
            return ref
        n, end = -count, at - count - 1
        if not ref["b"]:
            if at <= ref["a"][k] <= end:
                return None
            if ref["a"][k] > end:
                ref["a"][k] -= n
            return ref
        lo, hi = min(ref["a"][k], ref["b"][k]), max(ref["a"][k], ref["b"][k])
        if lo >= at and hi <= end:
            return None
        nlo = lo - n if lo > end else (at if lo >= at else lo)
        nhi = hi - n if hi > end else (at - 1 if hi >= at else hi)
        pa, pb = (ref["a"], ref["b"]) if ref["a"][k] <= ref["b"][k] else (ref["b"], ref["a"])
        pa[k], pb[k] = nlo, nhi
        return ref
    return rewrite_refs(formula, fn)
