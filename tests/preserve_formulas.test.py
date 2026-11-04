"""Files with functions this app cannot calculate: warn on open, write untouched formulas back unchanged.

    python3 tests/preserve_formulas.test.py      (needs Node.js for the calc worker)

The source file is written the way Excel stores it: prefixed formula text (_xlfn., _xlpm., _xludf.)
and a cached result in <v>. The test then opens it through the agent CLI, edits some cells, saves,
and compares the saved sheet XML cell by cell with the original.
"""
import io
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
XL = os.path.join(SK, "scripts", "xl.py")
sys.path.insert(0, os.path.join(SK, "scripts"))
from openpyxl import Workbook, load_workbook  # noqa: E402
from server.xlsx_io import inject_cached_values  # noqa: E402

fails = 0

# cell -> (formula as stored in the file, cached result Excel computed)
SRC = {
    "B1": ("=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A1)", 6),      # engine cannot parse a LAMBDA call
    "B2": ("=_xlfn.TEXTSPLIT(\"a,b\",\",\")", "a"),          # future function not in the engine
    "B3": ("=_xludf.MYFN(A1)", 42),                           # VBA user-defined function
    "B4": ("=_xlfn.NEWFUNC(A1)+1", 7),                        # prefixed function unknown to this app
    "B5": ("=_xlfn.NEWFUNC(A2)", 9),                          # same, will be edited
    "B6": ("=A1*10", 30),                                     # ordinary formula
    "B7": ("=_xlfn.XLOOKUP(3,A1:A2,A1:A2)", 3),               # supported future function
}


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


def xl(ws, *args):
    p = subprocess.run([sys.executable, XL, *args, "--root", ws], capture_output=True, text=True, encoding="utf-8")
    if p.returncode != 0:
        print(p.stdout, p.stderr)
    return p


def cells_xml(path):
    xml = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode("utf-8")
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', xml):
        f = re.search(r"<f[^>]*>(.*?)</f>", m.group(3))
        v = re.search(r"<v>(.*?)</v>", m.group(3))
        out[m.group(1)] = (f.group(1) if f else None, v.group(1) if v else None)
    return out


def unescape(s):
    return s.replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&") if s else s


def make(src):
    wb = Workbook()
    sh = wb.active
    sh.title = "S"
    sh["A1"], sh["A2"] = 3, 4
    for ref, (f, _v) in SRC.items():
        sh[ref] = f
    buf = io.BytesIO()
    wb.save(buf)
    with open(src, "wb") as fh:
        fh.write(inject_cached_values(buf.getvalue(), {"S": {k: v for k, (_f, v) in SRC.items()}}))


if len(sys.argv) == 3 and sys.argv[1] == "--make":  # used by tests/preserve.e2e.cjs
    make(sys.argv[2])
    sys.exit(0)
if len(sys.argv) == 4 and sys.argv[1] == "--cells":  # print saved <f>/<v> per cell as JSON
    print(json.dumps({"old": cells_xml(sys.argv[2]), "new": cells_xml(sys.argv[3])}))
    sys.exit(0)

from server.formula_compat import engine_functions  # noqa: E402
js_names = json.loads(subprocess.run(["node", "--input-type=module", "-e",
                                      "import('" + os.path.join(SK, "assets", "web", "js", "engine.js").replace("\\", "/")
                                      + "').then(m => console.log(JSON.stringify(m.FUNCTION_NAMES)))"],
                                     capture_output=True, text=True, check=True).stdout)
check("server's function list equals the engine's (" + str(len(js_names)) + ")", set(js_names) == set(engine_functions()),
      str(set(js_names) ^ set(engine_functions())))

ws = tempfile.mkdtemp(prefix="wx-keep-")
try:
    src = os.path.join(ws, "keep.xlsx")
    make(src)
    orig = cells_xml(src)

    xl(ws, "serve", "--port", "8813")
    out = xl(ws, "open", "keep.xlsx").stdout
    check("open warns about functions it cannot calculate",
          "계산할 수 없는 함수" in out and all(n in out for n in ("LAMBDA", "TEXTSPLIT", "MYFN (VBA)", "NEWFUNC")), out.strip())
    check("supported XLOOKUP is not reported", "XLOOKUP" not in out)
    d = json.loads(xl(ws, "describe", "--json").stdout)
    u = {x["cell"]: x for x in d["sheets"][0]["unsupported"]}
    check("describe lists the cells", sorted(u) == ["B1", "B2", "B3", "B4", "B5"], str(sorted(u)))
    r = json.loads(xl(ws, "read", "B1:B7", "--json").stdout)["rows"]
    check("unsupported cells show the file's last result", [x[0] for x in r] == [6, "a", 42, 7, 9, 30, 3], str(r))

    # edits: formula of B5 changed, B1 only reformatted, A1 value changed (B6 depends on it)
    xl(ws, "write", "B5", "=NEWFUNC(A2)*2")
    xl(ws, "format", "B1", "b=true")
    xl(ws, "write", "A1", "5")
    d = json.loads(xl(ws, "describe", "--json").stdout)
    u = {x["cell"]: x["kept"] for x in d["sheets"][0]["unsupported"]}
    check("describe: edited cell no longer keeps the original", u == {"B1": True, "B2": True, "B3": True, "B4": True, "B5": False}, str(u))

    xl(ws, "save", "--as", "out.xlsx")
    new = cells_xml(os.path.join(ws, "out.xlsx"))
    for ref in ("B1", "B2", "B3", "B4"):
        check(f"{ref} written back unchanged", new.get(ref) == orig.get(ref), f"{orig.get(ref)} -> {new.get(ref)}")
    check("edited B5 keeps the file's _xlfn. prefix, no #NAME? stored as its result",
          unescape(new["B5"][0]) == "_xlfn.NEWFUNC(A2)*2" and new["B5"][1] in (None, ""), str(new["B5"]))
    check("supported XLOOKUP recalculated with the new A1", new["B7"] == ("_xlfn.XLOOKUP(3,A1:A2,A1:A2)", "#N/A"), str(new["B7"]))
    check("ordinary formula recalculated (A1 3 -> 5)", new["B6"] == ("A1*10", "50"), str(new["B6"]))
    check("format change kept on the preserved cell", load_workbook(os.path.join(ws, "out.xlsx"))["S"]["B1"].font.b is True)
    xl(ws, "stop", "--force")
finally:
    subprocess.run([sys.executable, XL, "stop", "--force", "--root", ws], capture_output=True)
    shutil.rmtree(ws, ignore_errors=True)

print("ALL PASSED" if not fails else f"{fails} FAILED")
sys.exit(1 if fails else 0)
