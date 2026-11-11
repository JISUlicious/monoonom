"""The agent learns about the user's edits without asking: an edit's output says what the user changed
since the agent's last command, and warns when the edit replaced cells the user had just changed.
    python3 tests/user_changes.test.py
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XL = os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts", "xl.py")
fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra and not cond else ""))
    if not cond:
        fails += 1


ws = tempfile.mkdtemp(prefix="wx-uc-")


def xl(*a):
    p = subprocess.run([sys.executable, XL, *a, "--root", ws], capture_output=True, text=True)
    if p.returncode:
        raise SystemExit(f"xl {' '.join(a)} failed: {p.stderr}")
    return p.stdout


def preview(method, path, body=None):
    """A request the way the preview page sends it (no agent marker)."""
    st = json.load(open(os.path.join(ws, ".monoonom", "server.json")))
    req = urllib.request.Request(f"http://127.0.0.1:{st['port']}{path}", method=method,
                                 data=json.dumps(body).encode() if body is not None else None)
    req.add_header("X-Monoonom-Token", st["token"])
    req.add_header("X-Monoonom", "1")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read() or b"{}")


try:
    xl("serve", "--port", "8883")
    xl("new", "--name", "t.xlsx")
    doc = json.loads(xl("docs", "--json"))
    doc = (doc.get("docs") if isinstance(doc, dict) else doc)[0]
    did = doc["docId"]
    snap = preview("GET", f"/api/docs/{did}/snapshot")
    sid = (snap.get("model") or snap)["sheets"][0]["id"]

    def user_edit(cells, label="Typing"):
        preview("POST", f"/api/docs/{did}/ops", {"actor": "user", "clientId": "page-1", "label": label,
                                                  "ops": [{"op": "setCells", "sheet": sid, "cells": cells}]})

    out = xl("write", "A1", "x")
    check("no user edits: no note", "user changed" not in out and "warning" not in out, out)
    user_edit([[1, 1, {"v": 5}], [2, 1, {"v": 6}]])  # B2:B3
    out = xl("write", "D1", "y")
    check("the next edit says what the user changed (range, count)", "since your last command the user changed Typing: Sheet1!B2:B3 (2 cells)" in out, out)
    check("…no warning when the edit does not touch those cells", "warning" not in out, out)
    out = xl("write", "D2", "z")
    check("told once: the following edit has no note", "user changed" not in out, out)
    user_edit([[4, 0, {"v": 1}]])  # A5
    xl("read", "A1:D5")
    out = xl("write", "D3", "w")
    check("a read counts as seeing the document: no note after it", "user changed" not in out, out)
    user_edit([[6, 2, {"v": "mine"}]])  # C7
    out = xl("write", "C7", "agent")
    check("replacing a cell the user just changed: warning naming it and how to undo",
          "warning: this edit replaced cells the user had just changed: Sheet1!C7" in out and "xl.py undo" in out, out)
    for i in range(8):
        user_edit([[10 + i, 0, {"v": i}]], f"Edit {i}")
    out = xl("write", "H1", "k")
    check("many user edits: a few lines and where to see the rest", out.count("user changed") == 6 and "changes --since" in out, out)
    res = json.loads(subprocess.run([sys.executable, XL, "write", "H2", "j", "--json", "--root", ws], capture_output=True, text=True).stdout)
    check("--json output has no userChanges when nothing changed", "userChanges" not in res, json.dumps(res))
    # the preview's own edits through the agent API (actor user, e.g. a pivot chart) get no note
    user_edit([[20, 0, {"v": 1}]])
    r = preview("POST", f"/api/docs/{did}/edit", {"actor": "user", "edits": [{"op": "write", "range": "J1", "value": "p"}]})
    check("requests from the preview get no note", "userChanges" not in r, json.dumps(r))
finally:
    subprocess.run([sys.executable, XL, "stop", "--force", "--root", ws], capture_output=True)
    shutil.rmtree(ws, ignore_errors=True)

print("ALL PASSED" if not fails else f"{fails} FAILED")
sys.exit(1 if fails else 0)
