"""JS (browser) vs Python (server) formula reference rewriting on random formulas.

    python3 tests/refs_parity.test.py [N]
"""
import json
import os
import random
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts"))
from server import refs  # noqa: E402

N = int(sys.argv[1]) if len(sys.argv) > 1 else 10000
rnd = random.Random(20261104)
SHEETS = ["Sheet1", "데이터", "My Sheet", "O'Brien", "2024"]
FUNCS = ["SUM", "AVERAGE", "IF", "VLOOKUP", "INDEX", "MATCH", "LOG10", "ROUND", "XLOOKUP", "COUNTIFS"]


def col():
    return "".join(rnd.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ") for _ in range(rnd.choice([1, 1, 1, 2, 3])))[:3]


def cell():
    return f"{'$' if rnd.random() < .3 else ''}{col()}{'$' if rnd.random() < .3 else ''}{rnd.randint(1, 60)}"


def ref():
    k = rnd.random()
    if k < .5:
        r = cell()
    elif k < .75:
        r = f"{cell()}:{cell()}"
    elif k < .87:
        r = f"{'$' if rnd.random() < .3 else ''}{col()}:{'$' if rnd.random() < .3 else ''}{col()}"
    else:
        r = f"{rnd.randint(1, 60)}:{rnd.randint(1, 60)}"
    if rnd.random() < .3:
        sh = rnd.choice(SHEETS)
        q = refs.quote_sheet(sh) if rnd.random() < .8 else f"'{sh.replace(chr(39), chr(39) * 2)}'"
        r = f"{q}!{r}"
    return r


def expr(d=0):
    k = rnd.random()
    if d > 2 or k < .35:
        return ref()
    if k < .45:
        return rnd.choice(['"A1 text"', '"x""y"', "TRUE", "3.5", "1E+3", "#N/A", "ABC1D", "name_1"])
    if k < .75:
        return f"{rnd.choice(FUNCS)}({','.join(expr(d + 1) for _ in range(rnd.randint(1, 3)))})"
    return f"{expr(d + 1)}{rnd.choice(['+', '-', '*', '/', '&', '<>', '^', ' '])}{expr(d + 1)}"


cases = []
for _ in range(N):
    f = "=" + expr()
    k = rnd.random()
    if k < .4:
        cases.append({"f": f, "op": "translate", "args": [rnd.randint(-10, 10), rnd.randint(-5, 5)]})
    elif k < .9:
        cnt = rnd.choice([1, 2, 5, -1, -2, -5])
        cases.append({"f": f, "op": "adjust", "args": ["Sheet1", rnd.choice(SHEETS), rnd.choice(["row", "col"]), rnd.randint(0, 30), cnt]})
    else:
        cases.append({"f": f, "op": "rename", "args": [rnd.choice(SHEETS), "새 이름"]})

js = subprocess.run(["node", os.path.join(ROOT, "tests", "refs_parity.mjs")], input="\n".join(json.dumps(c, ensure_ascii=False) for c in cases),
                    capture_output=True, text=True, encoding="utf-8", check=True)
js_out = json.loads(js.stdout)


def py(c):
    if c["op"] == "translate":
        return refs.translate(c["f"], *c["args"])
    if c["op"] == "adjust":
        return refs.adjust(c["f"], *c["args"])
    old, new = c["args"]
    return refs.rewrite_refs(c["f"], lambda r: ({**r, "sheet": new} if r["sheet"] is not None and r["sheet"].lower() == old.lower() else r))


mism = [(c, j, py(c)) for c, j in zip(cases, js_out) if j != py(c)]
for c, j, p in mism[:10]:
    print("MISMATCH", json.dumps(c, ensure_ascii=False), "\n   js:", j, "\n   py:", p)
changed = sum(1 for c, j in zip(cases, js_out) if j != c["f"])
print(f"{N} formulas, {changed} rewritten, {len(mism)} mismatches")
sys.exit(1 if mism else 0)
