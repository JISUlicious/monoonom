"""CSV / TSV files: read into a one-sheet workbook model, and write one sheet's values back.

Reading: the encoding is tried as UTF-8 (with or without BOM), then CP949 (what Korean Excel writes
for "CSV"), then Latin-1; the delimiter is sniffed (`,` `;` tab `|`; `.tsv` is tab). Values are
read as typing them into Excel would: numbers, dates, percents and currency become numbers with a
format, `=` starts a formula, `'` keeps text.
Writing (only on an explicit `save --as name.csv`): the active sheet's values as displayed numbers
or text (formula results, dates as yyyy-mm-dd), UTF-8 with a BOM so Excel reads it right.
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import os
import re

ENCODINGS = ("utf-8-sig", "cp949", "latin-1")
DELIMS = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}


def _decode(data: bytes) -> tuple[str, str]:
    for enc in ENCODINGS:
        try:
            return data.decode(enc), ("utf-8" if enc == "utf-8-sig" else enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", "replace"), "latin-1"


def _delimiter(text: str, name: str) -> str:
    if name.lower().endswith(".tsv"):
        return "\t"
    try:
        return csv.Sniffer().sniff(text[:20000], delimiters="".join(DELIMS)).delimiter
    except csv.Error:
        return ","


def sheet_name(file_name: str) -> str:
    """Excel names the sheet after the file (without the extension), at most 31 characters."""
    stem = os.path.splitext(os.path.basename(file_name))[0] or "Sheet1"
    return re.sub(r"[\\/?*\[\]:]", "_", stem)[:31]


def read_csv(data: bytes, name: str) -> dict:
    from .agent_api import parse_value  # the same rules as the agent's `write` (typing in Excel)
    from .app import blank_workbook
    text, enc = _decode(data)
    delim = _delimiter(text, name)
    styles: list = []
    index: dict = {}
    cells = []
    for r, row in enumerate(csv.reader(io.StringIO(text), delimiter=delim)):
        for c, x in enumerate(row):
            out, nf = parse_value(x)
            if not out:
                continue
            if nf:
                if nf not in index:
                    index[nf] = len(styles)
                    styles.append({"nf": nf})
                out["s"] = index[nf]
            cells.append([r, c, out])
    model = blank_workbook(os.path.basename(name))
    sh = model["sheets"][0]
    sh.update({"name": sheet_name(name), "origName": None, "cells": cells})
    model["styles"] = styles
    model["warnings"] = [
        f"CSV ({enc}, {DELIMS.get(delim, repr(delim))}-separated) opened as one sheet. Saving writes "
        f"'{os.path.splitext(os.path.basename(name))[0]}.xlsx' (formulas and formatting kept); "
        f"saving as a .csv name writes the active sheet's values only."]
    return model


def _text(v, nf: str | None):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        f = (nf or "").lower()
        if re.search(r"[ymd]", re.sub(r'"[^"]*"|\[[^\]]*\]', "", f)) or re.search(r"h+:m", f):
            d = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))
            if re.search(r"h+:m", f) and not re.search(r"[ymd]", re.sub(r'"[^"]*"|\[[^\]]*\]', "", f)):
                return d.strftime("%H:%M:%S")
            return d.strftime("%Y-%m-%d %H:%M:%S") if d.hour or d.minute or d.second else d.strftime("%Y-%m-%d")
        if isinstance(v, float):
            v = float(f"{v:.15g}")
            return str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)
        return str(v)
    return str(v)


def write_csv(model: dict, path: str) -> tuple[bytes, list[str]]:
    """The active sheet as CSV (TSV for a .tsv path): (bytes, warnings)."""
    sheets = model.get("sheets") or []
    sm = sheets[min(model.get("activeSheet", 0), len(sheets) - 1)] if sheets else {"name": "Sheet1", "cells": []}
    styles = model.get("styles") or []
    grid: dict = {}
    has_formulas = False
    for r, c, it in sm.get("cells", []):
        st = styles[it["s"]] if isinstance(it.get("s"), int) and 0 <= it["s"] < len(styles) else {}
        v = it.get("cv") if it.get("f") else it.get("v")
        has_formulas = has_formulas or bool(it.get("f"))
        grid[(r, c)] = _text(v, (st or {}).get("nf"))
    buf = io.StringIO()
    w = csv.writer(buf, delimiter="\t" if path.lower().endswith(".tsv") else ",", lineterminator="\r\n")
    if grid:
        nr, nc = max(r for r, _ in grid) + 1, max(c for _, c in grid) + 1
        for r in range(nr):
            row = [grid.get((r, c), "") for c in range(nc)]
            while row and row[-1] == "":
                row.pop()
            w.writerow(row)
    lost = []
    if has_formulas:
        lost.append("formulas (their results are saved)")
    lost.append("formatting")
    if len(sheets) > 1:
        lost.append(f"the other {len(sheets) - 1} sheet(s)")
    if any(s.get("charts") or s.get("pivots") for s in sheets):
        lost.append("charts and pivot tables")
    warn = [f"Saved the values of sheet '{sm.get('name')}' as CSV (UTF-8); not in the file: {', '.join(lost)}. "
            "Save as .xlsx to keep them."]
    return buf.getvalue().encode("utf-8-sig"), warn
