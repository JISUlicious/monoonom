"""Formulas saved by web-excel open and calculate the same in another spreadsheet engine.

    python3 tests/excel_compat.test.py        (needs LibreOffice `soffice` and Node.js)

1. Build a workbook with ~150 formulas through the same save path the app uses (write_workbook).
2. Check the stored formula text (future functions need the _xlfn./_xlws. prefix in OOXML).
3. Let LibreOffice open + recalculate the file (headless convert), read its results with openpyxl
   and compare with our engine (Node worker = the browser's engine).
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SK = os.path.join(ROOT, ".agents", "skills", "web-excel")
sys.path.insert(0, os.path.join(SK, "scripts", "_vendor"))
sys.path.insert(0, os.path.join(SK, "scripts"))
from openpyxl import load_workbook  # noqa: E402
from server.app import write_workbook  # noqa: E402

DATA = [["이름", "수량", "단가", "날짜", "지역"],
        ["사과", 3, 1200, 46023, "서울"], ["배", 2, 2500, 46054, "부산"], ["포도", 5, 3000, 46082, "서울"],
        ["사과", 1, 1100, 46113, "대구"], ["귤", 7, 800, 46143, "부산"], ["감", 4, 1500, 46174, "서울"]]
# (formula, note) — no volatile functions (RAND, NOW, TODAY)
FORMULAS = [
    "=SUM(Data!B2:B7)", "=SUMIF(Data!A2:A7,\"사과\",Data!B2:B7)", "=SUMIFS(Data!B2:B7,Data!E2:E7,\"서울\",Data!C2:C7,\">1000\")",
    "=SUMPRODUCT(Data!B2:B7,Data!C2:C7)", "=SUMPRODUCT((Data!E2:E7=\"부산\")*Data!B2:B7)", "=PRODUCT(Data!B2:B4)", "=SUMSQ(1,2,3)",
    "=ROUND(2.345,2)", "=ROUNDUP(2.341,2)", "=ROUNDDOWN(-2.349,2)", "=MROUND(17,5)", "=INT(-3.5)", "=TRUNC(-3.7)", "=MOD(-7,3)",
    "=QUOTIENT(17,5)", "=CEILING(4.2,1)", "=FLOOR(4.8,2)", "=CEILING.MATH(-4.2)", "=FLOOR.MATH(4.8)", "=ABS(-5)", "=SIGN(-2)",
    "=POWER(2,10)", "=SQRT(16)", "=EXP(1)", "=LN(10)", "=LOG(1000)", "=LOG10(0.01)", "=PI()", "=FACT(6)", "=COMBIN(10,3)",
    "=PERMUT(5,2)", "=GCD(24,36)", "=LCM(4,6)", "=EVEN(3)", "=ODD(4)", "=SIN(1)", "=COS(1)", "=TAN(1)", "=ATAN2(1,2)",
    "=DEGREES(PI())", "=RADIANS(180)", "=SUBTOTAL(9,Data!B2:B7)", "=SUBTOTAL(101,Data!B2:B7)", "=AGGREGATE(9,6,Data!B2:B7)",
    "=AVERAGE(Data!B2:B7)", "=AVERAGEIF(Data!E2:E7,\"서울\",Data!B2:B7)", "=AVERAGEIFS(Data!C2:C7,Data!B2:B7,\">2\")",
    "=COUNT(Data!A1:E7)", "=COUNTA(Data!A1:E7)", "=COUNTBLANK(Data!A1:F7)", "=COUNTIF(Data!A2:A7,\"사*\")",
    "=COUNTIFS(Data!E2:E7,\"서울\",Data!B2:B7,\"<5\")", "=MAX(Data!C2:C7)", "=MIN(Data!C2:C7)", "=MAXIFS(Data!C2:C7,Data!E2:E7,\"부산\")",
    "=MINIFS(Data!C2:C7,Data!E2:E7,\"서울\")", "=MEDIAN(Data!B2:B7)", "=MODE(1,2,2,3)", "=MODE.SNGL(4,4,5)", "=STDEV(Data!B2:B7)",
    "=STDEV.S(Data!B2:B7)", "=STDEV.P(Data!B2:B7)", "=VAR(Data!B2:B7)", "=VAR.P(Data!B2:B7)", "=LARGE(Data!C2:C7,2)",
    "=SMALL(Data!C2:C7,2)", "=RANK(1500,Data!C2:C7)", "=RANK.EQ(3,Data!B2:B7,1)", "=RANK.AVG(3,{1,3,3,5})",
    "=PERCENTILE(Data!C2:C7,0.3)", "=PERCENTILE.INC(Data!C2:C7,0.3)", "=PERCENTILE.EXC(Data!C2:C7,0.3)", "=QUARTILE(Data!C2:C7,1)",
    "=QUARTILE.EXC(Data!C2:C7,1)", "=CORREL(Data!B2:B7,Data!C2:C7)", "=SLOPE(Data!C2:C7,Data!B2:B7)", "=INTERCEPT(Data!C2:C7,Data!B2:B7)",
    "=FORECAST(8,Data!C2:C7,Data!B2:B7)", "=FORECAST.LINEAR(8,Data!C2:C7,Data!B2:B7)", "=COVARIANCE.S(Data!B2:B7,Data!C2:C7)",
    "=GEOMEAN(2,8)", "=AVEDEV(2,4,6)", "=NORM.DIST(1,0,1,TRUE)", "=NORM.S.DIST(1,TRUE)",
    "=IF(SUM(Data!B2:B7)>20,\"많음\",\"적음\")", "=IFS(Data!B2>5,\"a\",Data!B2>2,\"b\",TRUE,\"c\")", "=IFERROR(1/0,\"오류\")",
    "=IFNA(NA(),\"없음\")", "=AND(TRUE,1>0)", "=OR(FALSE,0)", "=XOR(TRUE,TRUE)", "=NOT(FALSE)", "=SWITCH(2,1,\"one\",2,\"two\")",
    "=CHOOSE(2,\"a\",\"b\",\"c\")", "=VLOOKUP(\"포도\",Data!A2:C7,3,FALSE)", "=VLOOKUP(2000,Data!C2:C7,1,TRUE)",
    "=HLOOKUP(\"단가\",Data!A1:E7,3,FALSE)", "=XLOOKUP(\"귤\",Data!A2:A7,Data!C2:C7)", "=XLOOKUP(\"없음\",Data!A2:A7,Data!C2:C7,\"-\")",
    "=XMATCH(\"배\",Data!A2:A7)", "=INDEX(Data!A2:E7,3,2)", "=MATCH(1500,Data!C2:C7,0)", "=LOOKUP(2600,Data!C2:C4,Data!A2:A4)",
    "=SUM(OFFSET(Data!B1,1,0,3,1))", "=INDIRECT(\"Data!B3\")", "=ADDRESS(2,3)", "=ROW(Data!C5)", "=COLUMN(Data!C5)",
    "=ROWS(Data!A1:E7)", "=COLUMNS(Data!A1:E7)", "=SUM(Data!B2:INDEX(Data!B2:B7,3))",
    "=CONCATENATE(\"가\",\"나\",1)", "=CONCAT(Data!A2:A3)", "=TEXTJOIN(\",\",TRUE,Data!A2:A4)", "=LEFT(\"가나다라\",2)",
    "=RIGHT(\"abcdef\",3)", "=MID(\"abcdef\",2,3)", "=LEN(\"한글abc\")", "=UPPER(\"abc\")", "=LOWER(\"ABC\")", "=PROPER(\"hello world\")",
    "=TRIM(\"  a   b  \")", "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "=REPLACE(\"abcdef\",2,3,\"X\")", "=FIND(\"c\",\"abcabc\",4)",
    "=SEARCH(\"B*\",\"aabbcc\")", "=REPT(\"ab\",3)", "=EXACT(\"a\",\"A\")", "=CHAR(65)", "=CODE(\"A\")", "=VALUE(\"1234.5\")",
    "=NUMBERVALUE(\"1.234,5\",\",\",\".\")", "=TEXT(1234.567,\"#,##0.00\")", "=TEXT(46023,\"yyyy-mm-dd\")", "=TEXT(0.25,\"0%\")",
    "=FIXED(1234.567,1)", "=T(\"x\")", "=N(TRUE)",
    "=DATE(2026,2,30)", "=YEAR(46023)", "=MONTH(46023)", "=DAY(46023)", "=WEEKDAY(46023)", "=WEEKDAY(46023,2)", "=WEEKNUM(46023)",
    "=ISOWEEKNUM(46023)", "=EDATE(46023,1)", "=EOMONTH(46023,1)", "=DAYS(46100,46023)", "=DAYS360(46023,46100)",
    "=DATEDIF(43831,46023,\"Y\")", "=DATEDIF(43831,46023,\"YM\")", "=NETWORKDAYS(46023,46053)", "=WORKDAY(46023,10)",
    "=TIME(13,30,0)", "=HOUR(0.5625)", "=MINUTE(0.5625)", "=YEARFRAC(46023,46388,1)",
    "=PMT(0.05/12,12,1000000)", "=FV(0.04,10,-100)", "=PV(0.05,10,-1000)", "=NPV(0.1,-100,50,60)", "=IRR({-100,30,40,50})",
    "=NPER(0.01,-100,1000)", "=RATE(10,-100,800)",
    "=ISBLANK(Data!F1)", "=ISNUMBER(Data!B2)", "=ISTEXT(Data!A2)", "=ISERROR(1/0)", "=ISNA(NA())", "=ISEVEN(4)", "=ISODD(4)",
    "=ERROR.TYPE(1/0)", "=TYPE(\"a\")",
    "=-2^2", "=2^3^2", "=10%", "=0.1+0.2=0.3", "=1/0", "=LET(x,2,y,x*3,x+y)",
    "=NORM.S.DIST(-3.5,TRUE)", "=NORM.S.DIST(0.3,TRUE)", "=NORM.DIST(52,50,4,TRUE)", "=NORM.S.DIST(1.2,FALSE)",
    # day-count edge cases (month ends, Feb, leap years) for every basis / method
    *[f"=YEARFRAC({a},{b},{k})" for a, b in [("DATE(2024,1,31)", "DATE(2024,3,31)"), ("DATE(2023,2,28)", "DATE(2024,2,29)"),
                                            ("DATE(2024,2,29)", "DATE(2025,2,28)"), ("DATE(2023,3,1)", "DATE(2024,2,29)"),
                                            ("DATE(2020,6,15)", "DATE(2026,1,10)"), ("DATE(2025,12,31)", "DATE(2025,1,1)"),
                                            ("DATE(2023,1,31)", "DATE(2023,2,28)")] for k in range(5)],
    *[f"=DAYS360({a},{b}{m})" for a, b in [("DATE(2023,2,28)", "DATE(2023,3,31)"), ("DATE(2024,1,31)", "DATE(2024,3,31)"),
                                          ("DATE(2024,1,30)", "DATE(2024,3,31)"), ("DATE(2024,2,29)", "DATE(2024,3,31)"),
                                          ("DATE(2024,3,31)", "DATE(2024,1,15)")] for m in ("", ",TRUE")],
]
# Known differences that are LibreOffice's, not ours (checked against Microsoft's function reference):
LO_KNOWN = {
    "=LET": "LibreOffice < 24.8 has no LET",
    "=XLOOKUP": "LibreOffice < 24.8 has no XLOOKUP/XMATCH",
    "=XMATCH": "LibreOffice < 24.8 has no XLOOKUP/XMATCH",
}


def same(mine, theirs):
    """Our engine value vs a value read back with openpyxl (dates/times come back as datetime)."""
    import datetime as dt
    mine = mine["err"] if isinstance(mine, dict) else mine
    if isinstance(theirs, (dt.datetime, dt.date, dt.time)):
        if isinstance(theirs, dt.time):
            theirs = (theirs.hour * 3600 + theirs.minute * 60 + theirs.second + theirs.microsecond / 1e6) / 86400
        else:
            if not isinstance(theirs, dt.datetime):
                theirs = dt.datetime(theirs.year, theirs.month, theirs.day)
            theirs = (theirs - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(mine, bool) or isinstance(theirs, bool):
        return mine == theirs
    if isinstance(mine, (int, float)) and isinstance(theirs, (int, float)):
        return abs(mine - theirs) <= 1e-12 * max(1, abs(theirs))
    return mine == theirs or (mine is not None and str(mine) == str(theirs))


def main():
    tmp = tempfile.mkdtemp(prefix="wx-compat-")
    try:
        model = {"name": "c.xlsx", "format": "xlsx", "activeSheet": 0, "styles": [], "definedNames": {},
                 "defaultFont": {"name": "Calibri", "size": 11},
                 "sheets": [{"id": "s1", "name": "T", "cells": [], "cols": {}, "rows": {}, "merges": [], "charts": []},
                            {"id": "s2", "name": "Data", "cells": [[r, c, {"v": v}] for r, row in enumerate(DATA) for c, v in enumerate(row)],
                             "cols": {}, "rows": {}, "merges": [], "charts": []}]}
        for i, f in enumerate(FORMULAS):
            model["sheets"][0]["cells"].append([i, 0, {"f": f}])
        # our engine (same code as the browser) -> expected values
        node = subprocess.run(["node", os.path.join(SK, "scripts", "calc", "worker.mjs")], input=json.dumps(model, ensure_ascii=False),
                              capture_output=True, text=True, encoding="utf-8", check=True)
        ours = {r: v for sid, r, c, v in json.loads(node.stdout)["values"] if sid == "s1"}
        for r, c, item in model["sheets"][0]["cells"]:  # what DocSession.file_model() does on save
            v = ours.get(r)
            item["cv"] = v["err"] if isinstance(v, dict) else v
        path = os.path.join(tmp, "c.xlsx")
        write_workbook(path, "xlsx", model, None)

        # 1) stored formula text
        xml = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode("utf-8")
        stored = re.findall(r"<f[^>]*>(.*?)</f>", xml)
        need = {"MAXIFS", "MINIFS", "IFS", "SWITCH", "XOR", "IFNA", "XLOOKUP", "XMATCH", "CONCAT", "TEXTJOIN", "STDEV.S",
                "STDEV.P", "VAR.P", "MODE.SNGL", "RANK.EQ", "RANK.AVG", "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.EXC",
                "CEILING.MATH", "FLOOR.MATH", "AGGREGATE", "COVARIANCE.S", "FORECAST.LINEAR", "NORM.DIST", "NORM.S.DIST",
                "ISOWEEKNUM", "DAYS", "NUMBERVALUE", "LET"}
        missing = [n for n in need if not any(f"_xlfn.{n}(" in s for s in stored)]
        if not any("_xlfn.LET(_xlpm.x," in s for s in stored):
            missing.append("_xlpm. LET names")
        print(f"{'PASS' if not missing else 'FAIL'} future functions stored with _xlfn. prefix ({len(need) - len(missing)}/{len(need)})"
              + (f" — missing: {missing}" if missing else ""))
        cached = re.findall(r"<f[ >](?:(?!</c>).)*?</f><v>[^<]+</v>", xml)
        print(f"{'PASS' if len(cached) == len(FORMULAS) else 'FAIL'} formula cells carry a cached result: {len(cached)}/{len(FORMULAS)}"
              f" (fullCalcOnLoad={'yes' if 'fullCalcOnLoad=\"1\"' in zipfile.ZipFile(path).read('xl/workbook.xml').decode() else 'no'})")
        ro = load_workbook(path, data_only=True)["T"]
        miss = [f for i, f in enumerate(FORMULAS) if not same(ours.get(i), ro.cell(i + 1, 1).value)]
        print(f"{'PASS' if not miss else 'FAIL'} a reader that does not calculate (openpyxl data_only) sees our results: "
              f"{len(FORMULAS) - len(miss)}/{len(FORMULAS)}" + (f" — {miss[:5]}" if miss else ""))
        missing += [f"cached {len(FORMULAS) - len(cached)}"] if len(cached) != len(FORMULAS) else []
        missing += miss

        # 2) LibreOffice opens + recalculates. LO reuses cached xlsx values by default, so give it a copy
        #    without them (otherwise it would only echo our numbers back).
        nocache = os.path.join(tmp, "n", "c.xlsx")
        os.makedirs(os.path.dirname(nocache))
        for _r, _c, item in model["sheets"][0]["cells"]:
            item.pop("cv", None)
        write_workbook(nocache, "xlsx", model, None)
        out = os.path.join(tmp, "lo")
        lo = subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, nocache],
                            capture_output=True, text=True, timeout=180)
        conv = os.path.join(out, "c.xlsx")
        if not os.path.exists(conv):
            print("FAIL LibreOffice could not open the file:", lo.stdout, lo.stderr)
            return 1
        ws = load_workbook(conv, data_only=True)["T"]
        wsf = load_workbook(conv)["T"]
        bad, err_lo, known = [], [], []
        for i, f in enumerate(FORMULAS):
            theirs = ws.cell(i + 1, 1).value
            ok = same(ours.get(i), theirs)
            if not ok and any(f.startswith(k) for k in LO_KNOWN):
                known.append((f, LO_KNOWN[next(k for k in LO_KNOWN if f.startswith(k))]))
            elif isinstance(theirs, str) and theirs.startswith("#") and not (isinstance(ours.get(i), dict)):
                err_lo.append((f, theirs, wsf.cell(i + 1, 1).value))
            elif not ok:
                bad.append((f, mine, theirs))
        n = len(FORMULAS)
        print(f"{'PASS' if not bad and not err_lo else 'FAIL'} LibreOffice recalculation matches our engine: "
              f"{n - len(bad) - len(err_lo) - len(known)}/{n}, {len(known)} known LibreOffice gaps")
        for f, why in known:
            print(f"   LO gap    {f:55s} ({why})")
        for f, theirs, kept in err_lo:
            print(f"   LO error  {f:55s} -> {theirs}   (formula after round trip: {kept})")
        for f, mine, theirs in bad:
            print(f"   differs   {f:55s} ours={mine!r} LO={theirs!r}")
        return 1 if (missing or bad or err_lo) else 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
