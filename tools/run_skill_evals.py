"""Run the web-excel skill's evaluations with the Claude Code CLI (`claude -p`).

    python3 tools/run_skill_evals.py trigger   [--models haiku,sonnet,opus] [--out DIR]
    python3 tools/run_skill_evals.py scenarios [--models haiku,sonnet,opus] [--only ID,...] [--out DIR]
    python3 tools/run_skill_evals.py report    --out DIR

Each run gets a fresh workspace with the skill installed as Claude Code reads it
(.claude/skills/web-excel: the generated SKILL.md plus links into the package) and a copy of
samples/sample.xlsx (also under samples/).

trigger:   evals/trigger_queries.json. The agent may look around (Glob, Read) but Bash is disallowed,
           so nothing runs; a run "triggers" when it calls the Skill tool for web-excel within 3 turns.
scenarios: evals/evals.json. A server is started for the workspace and the preview is open in a
           headless browser for the whole run (the "user"); some scenarios first act as the user
           (select a range, edit cells). The agent then works without limits on Bash. `Skill` must be in
           --allowedTools: the Claude Code SKILL.md has allowed-tools, which makes invoking the skill a
           permission request, denied in a -p run otherwise ("permission_denied ... Execute skill"). A grader
           (claude -p, opus) judges each expectation from the transcript and the workbook state.
Results: DIR/trigger.json, DIR/scenarios.json, DIR/transcripts/*.jsonl; `report` prints a summary.
"""
import argparse
import concurrent.futures as cf
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, ".agents", "skills", "web-excel")
CC_SKILL = os.path.join(ROOT, ".claude", "skills", "web-excel", "SKILL.md")
SAMPLE = os.path.join(ROOT, "samples", "sample.xlsx")


def workspace(tag):
    ws = tempfile.mkdtemp(prefix=f"wx-eval-{tag}-")
    sk = os.path.join(ws, ".claude", "skills", "web-excel")
    os.makedirs(sk)
    shutil.copy(CC_SKILL, os.path.join(sk, "SKILL.md"))
    for e in ("scripts", "references", "assets", "LICENSE.txt", "THIRD_PARTY_NOTICES.md"):
        os.symlink(os.path.join(PKG, e), os.path.join(sk, e))
    os.makedirs(os.path.join(ws, "samples"))
    shutil.copy(SAMPLE, os.path.join(ws, "sample.xlsx"))
    shutil.copy(SAMPLE, os.path.join(ws, "samples", "sample.xlsx"))
    return ws


def claude(prompt, cwd, model, extra, timeout, log):
    cmd = ["claude", "-p", prompt, "--model", model, "--output-format", "stream-json", "--verbose",
           "--no-session-persistence", *extra]
    t0 = time.time()
    try:
        p = subprocess.run(cmd, cwd=cwd, stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=timeout)
        out = p.stdout
    except subprocess.TimeoutExpired as e:
        out = (e.stdout or b"").decode() if isinstance(e.stdout, bytes) else (e.stdout or "")
    with open(log, "w", encoding="utf-8") as f:
        f.write(out)
    events = []
    for line in out.splitlines():
        try:
            events.append(json.loads(line))
        except ValueError:
            pass
    return events, time.time() - t0


def tool_uses(events):
    for o in events:
        if o.get("type") == "assistant":
            for c in o["message"]["content"]:
                if c.get("type") == "tool_use":
                    yield c


def result_of(events):
    r = next((o for o in reversed(events) if o.get("type") == "result"), {})
    return {"cost": r.get("total_cost_usd"), "turns": r.get("num_turns"), "subtype": r.get("subtype"), "text": r.get("result")}


# ------------------------------------------------------------------ trigger

def run_trigger(models, out):
    tq = json.load(open(os.path.join(PKG, "evals", "trigger_queries.json"), encoding="utf-8"))
    jobs = [(q, True, m) for q in tq["should_trigger"] for m in models] + [(q, False, m) for q in tq["should_not_trigger"] for m in models]
    os.makedirs(os.path.join(out, "transcripts"), exist_ok=True)

    def one(job):
        q, want, m = job
        ws = workspace("trig")
        log = os.path.join(out, "transcripts", f"trigger-{m}-{abs(hash(q)) % 10**8}.jsonl")
        ev, secs = claude(q, ws, m, ["--max-turns", "3", "--disallowedTools", "Bash", "--max-budget-usd", "0.5"], 300, log)
        uses = list(tool_uses(ev))
        names = [u["name"] + (f"({u['input'].get('skill')})" if u["name"] == "Skill" else "") for u in uses]
        fired = any(u["name"] == "Skill" and "web-excel" in json.dumps(u["input"]) for u in uses)
        other = [u["input"].get("skill") for u in uses if u["name"] == "Skill" and "web-excel" not in json.dumps(u["input"])]
        shutil.rmtree(ws, ignore_errors=True)
        return {"query": q, "should_trigger": want, "model": m, "triggered": fired, "other_skills": other,
                "tools": names, "ok": fired == want, "secs": round(secs), **{k: v for k, v in result_of(ev).items() if k != "text"}}

    with cf.ThreadPoolExecutor(6) as ex:
        res = list(ex.map(one, jobs))
    json.dump(res, open(os.path.join(out, "trigger.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return res


# ------------------------------------------------------------------ scenarios

PREP = {
    # the "user" selects 매출!B2:C4 in the preview
    "react-to-user-selection": "async (p) => { await p.evaluate(() => { const a = window.webExcel; a.activateSheet(0); a.selectRange({ r1: 1, c1: 1, r2: 3, c2: 2 }); }); }",
    # the "user" edits two cells in the preview
    "user-edits-meanwhile": "async (p) => { await p.evaluate(() => { const a = window.webExcel; a.activateSheet(0); a.select(2, 1); a.focusGrid(); }); "
                            "await p.keyboard.type('999'); await p.keyboard.press('Enter'); "
                            "await p.evaluate(() => { const a = window.webExcel; a.select(4, 3); a.focusGrid(); }); "
                            "await p.keyboard.type('77'); await p.keyboard.press('Enter'); }",
}

BROWSER = r"""
const { chromium } = require('playwright');
const [url, prep] = [process.argv[2], process.argv[3]];
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1300, height: 800 } });
  await p.goto(url); await p.waitForFunction(() => window.webExcel?.doc?.sync, null, { timeout: 20000 });
  await p.waitForTimeout(500);
  if (prep) await eval(prep)(p);
  await p.waitForTimeout(800);
  console.log('READY');
  process.stdin.on('data', async () => { await b.close(); process.exit(0); });
  setTimeout(async () => { await b.close(); process.exit(0); }, 1800000);
})().catch((e) => { console.error(e); process.exit(1); });
"""


def xl(ws, *a):
    return subprocess.run([sys.executable, os.path.join(ws, ".claude", "skills", "web-excel", "scripts", "xl.py"), *a, "--root", ws],
                          capture_output=True, text=True, timeout=120).stdout


def run_scenarios(models, out, only=None):
    ev = json.load(open(os.path.join(PKG, "evals", "evals.json"), encoding="utf-8"))["evals"]
    if only:
        ev = [e for e in ev if e["id"] in only]
    jobs = [(e, m, i) for i, (e, m) in enumerate((e, m) for e in ev for m in models)]
    os.makedirs(os.path.join(out, "transcripts"), exist_ok=True)
    bjs = os.path.join(out, "browser.cjs")
    open(bjs, "w").write(BROWSER)

    def one(job):
        e, m, i = job
        ws = workspace(e["id"][:12])
        port = 8900 + i
        xl(ws, "serve", "--port", str(port))
        xl(ws, "open", "samples/sample.xlsx")
        url = xl(ws, "url", "--bare").strip()
        env = dict(os.environ, NODE_PATH=subprocess.run(["npm", "root", "-g"], capture_output=True, text=True).stdout.strip())
        br = subprocess.Popen(["node", bjs, url, PREP.get(e["id"], "")], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True, env=env)
        ready = br.stdout.readline().strip() == "READY"
        log = os.path.join(out, "transcripts", f"scenario-{e['id']}-{m}.jsonl")
        evs, secs = claude(e["prompt"], ws, m, ["--max-turns", "40", "--allowedTools", "Skill Bash Read Glob Grep Write Edit", "--max-budget-usd", "3"], 1200, log)
        # workbook state for the grader
        state = {
            "describe": xl(ws, "describe"),
            "매출 formulas": xl(ws, "read", "매출!A1:H10", "--what", "formulas"),
            "매출 formats": xl(ws, "read", "매출!A1:H8", "--what", "all")[:4000],
            "요약 sheet": xl(ws, "read", "요약!A1:F15", "--what", "formulas") if "요약" in xl(ws, "describe") else "(no 요약 sheet)",
            "saved (backups)": os.listdir(os.path.join(ws, ".web-excel", "backups")) if os.path.isdir(os.path.join(ws, ".web-excel", "backups")) else [],
            "files": sorted(os.listdir(ws)),
        }
        try:
            br.stdin.write("x\n"); br.stdin.flush()
        except Exception:
            pass
        xl(ws, "stop", "--force")
        steps = []
        results = {}
        for o in evs:
            if o.get("type") == "user":
                for c in o["message"]["content"] if isinstance(o["message"]["content"], list) else []:
                    if c.get("type") == "tool_result":
                        txt = c.get("content")
                        txt = txt if isinstance(txt, str) else json.dumps(txt, ensure_ascii=False)
                        results[c["tool_use_id"]] = txt[:600]
        for u in tool_uses(evs):
            steps.append({"tool": u["name"], "input": u["input"].get("command") or u["input"].get("skill") or json.dumps(u["input"], ensure_ascii=False)[:300], "output": results.get(u["id"], "")})
        r = result_of(evs)
        grade = grade_run(e, steps, r.get("text") or "", state)
        shutil.rmtree(ws, ignore_errors=True)
        return {"id": e["id"], "model": m, "browser_ready": ready, "skill_used": any(s["tool"] == "Skill" and "web-excel" in s["input"] for s in steps),
                "steps": len(steps), "cost": r["cost"], "turns": r["turns"], "subtype": r["subtype"], "secs": round(secs),
                "grade": grade, "passed": sum(1 for g in grade if g.get("pass")), "total": len(e["expectations"])}

    with cf.ThreadPoolExecutor(4) as ex:
        res = list(ex.map(one, jobs))
    json.dump(res, open(os.path.join(out, "scenarios.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return res


def grade_run(e, steps, final, state):
    lines = [f"[{i + 1}] {s['tool']}: {s['input']}\n    → {s['output'][:400]}" for i, s in enumerate(steps)]
    prompt = (
        "You grade one run of an AI agent that used the web-excel skill (a CLI, xl.py, editing an Excel workbook "
        "the user sees in a live preview). Judge each expectation strictly from the evidence below; if the evidence "
        "does not show it, it fails.\n\n"
        f"USER REQUEST:\n{e['prompt']}\n\nEXPECTED OUTCOME:\n{e['expected_output']}\n\n"
        "EXPECTATIONS:\n" + "\n".join(f"{i + 1}. {x}" for i, x in enumerate(e["expectations"])) +
        "\n\nTOOL CALLS (in order, with truncated output):\n" + "\n".join(lines)[:60000] +
        f"\n\nAGENT'S FINAL MESSAGE TO THE USER:\n{final[:6000]}\n\nWORKBOOK STATE AFTER THE RUN:\n" + json.dumps(state, ensure_ascii=False)[:20000] +
        '\n\nAnswer with only a JSON array, one object per expectation in order: [{"expectation": "...", "pass": true|false, "evidence": "short quote or reason"}]'
    )
    p = subprocess.run(["claude", "-p", prompt, "--model", "opus", "--output-format", "json", "--max-turns", "1",
                        "--tools", "", "--no-session-persistence"], stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=600)
    try:
        text = json.loads(p.stdout)["result"]
        return json.loads(re.search(r"\[.*\]", text, re.S).group(0))
    except Exception as exc:  # keep the run; mark grading as failed
        return [{"expectation": x, "pass": False, "evidence": f"grader error: {exc}: {p.stdout[:200]} {p.stderr[:200]}"} for x in e["expectations"]]


# ------------------------------------------------------------------ report

def report(out):
    lines = []
    tp = os.path.join(out, "trigger.json")
    if os.path.exists(tp):
        t = json.load(open(tp, encoding="utf-8"))
        models = sorted({x["model"] for x in t}, key=["haiku", "sonnet", "opus"].index)
        lines.append("## Trigger\n\n| model | should trigger (fired) | should not (stayed quiet) | accuracy | cost |\n|---|---|---|---|---|")
        for m in models:
            r = [x for x in t if x["model"] == m]
            pos = [x for x in r if x["should_trigger"]]; neg = [x for x in r if not x["should_trigger"]]
            lines.append(f"| {m} | {sum(x['triggered'] for x in pos)}/{len(pos)} | {sum(not x['triggered'] for x in neg)}/{len(neg)} | "
                         f"{sum(x['ok'] for x in r)}/{len(r)} | ${sum(x['cost'] or 0 for x in r):.2f} |")
        miss = [x for x in t if not x["ok"]]
        if miss:
            lines.append("\nMisses:\n" + "\n".join(f"- {x['model']}: {'missed' if x['should_trigger'] else 'false trigger'} — {x['query']} (tools: {', '.join(x['tools']) or '-'})" for x in miss))
    sp = os.path.join(out, "scenarios.json")
    if os.path.exists(sp):
        s = json.load(open(sp, encoding="utf-8"))
        models = sorted({x["model"] for x in s}, key=["haiku", "sonnet", "opus"].index)
        ids = list(dict.fromkeys(x["id"] for x in s))
        lines.append("\n## Scenarios (expectations passed)\n\n| scenario | " + " | ".join(models) + " |\n|---|" + "---|" * len(models))
        for i in ids:
            row = []
            for m in models:
                x = next((y for y in s if y["id"] == i and y["model"] == m), None)
                row.append(f"{x['passed']}/{x['total']}" + ("" if x["skill_used"] else " (no skill)") if x else "-")
            lines.append(f"| {i} | " + " | ".join(row) + " |")
        for m in models:
            r = [x for x in s if x["model"] == m]
            lines.append(f"\n{m}: {sum(x['passed'] for x in r)}/{sum(x['total'] for x in r)} expectations, ${sum(x['cost'] or 0 for x in r):.2f}")
        fails = [(x, g) for x in s for g in x["grade"] if not g.get("pass")]
        if fails:
            lines.append("\nFailed expectations:\n" + "\n".join(f"- {x['id']} / {x['model']}: {g['expectation']} — {g.get('evidence', '')}" for x, g in fails))
    print("\n".join(lines))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("what", choices=["trigger", "scenarios", "report"])
    ap.add_argument("--models", default="haiku,sonnet,opus")
    ap.add_argument("--only")
    ap.add_argument("--out", default=os.path.join(tempfile.gettempdir(), "wx-evals"))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    models = a.models.split(",")
    if a.what == "trigger":
        run_trigger(models, a.out)
    elif a.what == "scenarios":
        run_scenarios(models, a.out, a.only.split(",") if a.only else None)
    report(a.out)
