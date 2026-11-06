"""Formula text compatibility between the UI and the file format.

Functions added after Excel 2007 are stored in OOXML with an ``_xlfn.`` prefix
(dynamic-array ones also ``_xlws.``); Excel shows #NAME? if the prefix is missing.
See [MS-XLSX] §2.2.2 "Formulas" and Microsoft's "_xlfn. prefix" support note.
"""
from __future__ import annotations

import functools
import os
import re

FUTURE_FUNCTIONS = {
    "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV", "BITAND", "BITOR",
    "BITXOR", "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.INV", "CONCAT",
    "CONFIDENCE.NORM", "CONFIDENCE.T", "COVARIANCE.P", "COVARIANCE.S", "DAYS", "ERF.PRECISE",
    "EXPON.DIST", "F.DIST", "F.INV", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.LINEAR",
    "GAMMA.DIST", "GAMMA.INV", "IFNA", "IFS", "ISOWEEKNUM", "LET", "LOGNORM.DIST",
    "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "NETWORKDAYS.INTL", "NORM.DIST",
    "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PERCENTILE.EXC",
    "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "POISSON.DIST", "QUARTILE.EXC",
    "QUARTILE.INC", "RANK.AVG", "RANK.EQ", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST",
    "T.INV", "TEXTJOIN", "VAR.P", "VAR.S", "WEIBULL.DIST", "WORKDAY.INTL", "XLOOKUP",
    "XMATCH", "XOR", "SEQUENCE", "RANDARRAY", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "VSTACK", "HSTACK", "TOCOL", "TOROW", "CHOOSECOLS", "CHOOSEROWS", "TAKE", "DROP",
    # Excel 2013 additions
    "FORMULATEXT", "ISFORMULA", "SHEET", "SHEETS", "UNICHAR", "UNICODE", "ACOT", "ACOTH", "ARABIC", "BASE",
    "COMBINA", "COT", "COTH", "CSC", "CSCH", "DECIMAL", "GAMMA", "GAUSS", "PDURATION", "PERMUTATIONA", "PHI",
    "RRI", "SEC", "SECH", "SKEW.P", "LAMBDA",
}
# dynamic-array functions additionally carry the _xlws. namespace
XLWS_FUNCTIONS = {"FILTER", "SORT", "SORTBY", "UNIQUE"}

_STR_OR_SHEET = re.compile(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')')
_FN = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*\(")


def _map_outside_literals(formula: str, fn) -> str:
    parts = _STR_OR_SHEET.split(formula)
    return "".join(p if i % 2 else fn(p) for i, p in enumerate(parts))


def strip_prefixes(formula: str) -> str:
    return _map_outside_literals(formula, lambda s: re.sub(r"_xlfn\.(?:_xlws\.)?|_xlws\.|_xlpm\.", "", s))


def _let_names(formula: str):
    """Parameter names declared by LET(name1, value1, ..., calc) calls (outside string literals)."""
    names = set()
    code = _map_outside_literals(formula, lambda s: s)
    for m in re.finditer(r"(?<![A-Za-z0-9_.])LET\s*\(", code, re.I):
        depth, arg, i, start = 0, 0, m.end(), m.end()
        args = []
        in_str = False
        while i < len(code):
            ch = code[i]
            if ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch in "({":
                    depth += 1
                elif ch in ")}":
                    if depth == 0:
                        args.append(code[start:i])
                        break
                    depth -= 1
                elif ch == "," and depth == 0:
                    args.append(code[start:i])
                    start = i + 1
            i += 1
        for a in args[:-1][::2]:
            a = a.strip()
            if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", a):
                names.add(a.lower())
    return names


def add_prefixes(formula: str, extra: dict | None = None) -> str:
    """Future functions get _xlfn./_xlws.; LET parameter names get _xlpm. (OOXML storage form).

    ``extra`` maps more function names to the prefix the opened file used for them (see
    prefix_map), so functions this module does not know keep their prefix after an edit."""
    params = _let_names(formula)
    if params:
        pat = re.compile(r"(?<![A-Za-z0-9_.$!'])([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!])")
        formula = _map_outside_literals(formula, lambda s: pat.sub(lambda m: "_xlpm." + m.group(1) if m.group(1).lower() in params else m.group(1), s))

    def repl(m: re.Match) -> str:
        name = m.group(1).upper()
        if name in XLWS_FUNCTIONS:
            return "_xlfn._xlws." + m.group(0)
        if name in FUTURE_FUNCTIONS:
            return "_xlfn." + m.group(0)
        if extra and name in extra:
            return extra[name] + m.group(0)
        return m.group(0)

    return _map_outside_literals(formula, lambda s: _FN.sub(repl, s))


def prefix_map(raw: str) -> dict:
    """{FUNCTION: '_xlfn.' | '_xlfn._xlws.'} for every prefixed call in a formula as stored in a file."""
    out = {}

    def scan(part):
        for m in re.finditer(r"(_xlfn\.(?:_xlws\.)?)([A-Za-z][A-Za-z0-9_.]*)\s*\(", part):
            out.setdefault(m.group(2).upper(), m.group(1))
        return part
    _map_outside_literals(raw, scan)
    return out


@functools.lru_cache(maxsize=1)
def engine_functions() -> frozenset:
    """Function names the browser/Node engine implements, read from engine.js's def(...) table.
    tests/preserve_formulas.test.py checks this equals the engine's own FUNCTION_NAMES."""
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                        "assets", "web", "js", "engine.js")
    try:
        with open(path, encoding="utf-8") as f:
            src = f.read()
    except OSError:
        return frozenset()
    names = set()
    for m in re.finditer(r"^def\('([^']+)'", src, re.M):
        names.update(m.group(1).split())
    # table-driven definitions: for (const [n, f] of Object.entries({ SIN: Math.sin, ... })) def(n, ...)
    for m in re.finditer(r"^for \(const \[n,[^\n]*?Object\.entries\(\{([^}]*)\}", src, re.M):
        names.update(re.findall(r"([A-Z][A-Z0-9.]*)\s*:", m.group(1)))
    return frozenset(names)


_CALL = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")


@functools.lru_cache(maxsize=4096)
def unsupported_functions(formula: str) -> tuple:
    """Functions called in a (prefix-free) formula that the engine cannot calculate, e.g. LAMBDA,
    TEXTSPLIT, user-defined VBA (_xludf.) or add-in functions. Empty when everything is supported."""
    known = engine_functions()
    if not known:
        return ()
    found = []

    def scan(part):
        found.extend(m.group(1) for m in _CALL.finditer(part))
        return part
    _map_outside_literals(formula, scan)
    out = []
    for name in found:
        u = name.upper()
        if u.startswith("_XLUDF."):
            u = name[7:].upper() + " (VBA)"
        elif u.startswith("_XLL."):
            u = name[5:].upper() + " (add-in)"
        elif u in known:
            continue
        if u not in out:
            out.append(u)
    return tuple(out)
