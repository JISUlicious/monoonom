"""Formula text compatibility between the UI and the file format.

Functions added after Excel 2007 are stored in OOXML with an ``_xlfn.`` prefix
(dynamic-array ones also ``_xlws.``); Excel shows #NAME? if the prefix is missing.
See [MS-XLSX] §2.2.2 "Formulas" and Microsoft's "_xlfn. prefix" support note.
"""
from __future__ import annotations

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
}
# dynamic-array functions additionally carry the _xlws. namespace
XLWS_FUNCTIONS = {"FILTER", "SORT", "SORTBY", "UNIQUE"}

_STR_OR_SHEET = re.compile(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')')
_FN = re.compile(r"(?<![A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)\s*\(")


def _map_outside_literals(formula: str, fn) -> str:
    parts = _STR_OR_SHEET.split(formula)
    return "".join(p if i % 2 else fn(p) for i, p in enumerate(parts))


def strip_prefixes(formula: str) -> str:
    return _map_outside_literals(formula, lambda s: re.sub(r"_xlfn\.(?:_xlws\.)?|_xlws\.", "", s))


def add_prefixes(formula: str) -> str:
    def repl(m: re.Match) -> str:
        name = m.group(1).upper()
        if name in XLWS_FUNCTIONS:
            return "_xlfn._xlws." + m.group(0)
        if name in FUTURE_FUNCTIONS:
            return "_xlfn." + m.group(0)
        return m.group(0)

    return _map_outside_literals(formula, lambda s: _FN.sub(repl, s))
