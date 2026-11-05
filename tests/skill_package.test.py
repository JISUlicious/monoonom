"""Skill package checks: python3 tests/skill_package.test.py

- the Claude Code SKILL files are rebuilt from SKILL.md (tools/build_claude_code_skill.py --check)
- .claude/skills/monoonom is the Claude Code skill: generated SKILL.md + links into the package
- frontmatter within the Agent Skills spec limits (https://agentskills.io/specification); the full
  check is the official validator: skills-ref validate <dir>
"""
import os
import re
import subprocess
import sys

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, ".agents", "skills", "monoonom")
CC = os.path.join(ROOT, ".claude", "skills", "monoonom")
fails = 0


def check(name, cond, extra=""):
    global fails
    print(("PASS " if cond else "FAIL ") + name + (f" — {extra}" if extra else ""))
    if not cond:
        fails += 1


p = subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build_claude_code_skill.py"), "--check"], capture_output=True, text=True)
check("SKILL.claude-code.md and .claude/skills/monoonom/SKILL.md are up to date with SKILL.md", p.returncode == 0, p.stdout.strip())

check(".claude/skills/monoonom is a folder (its SKILL.md is the Claude Code version)", os.path.isdir(CC) and not os.path.islink(CC))
check("…its SKILL.md is a file, not a link", os.path.isfile(os.path.join(CC, "SKILL.md")) and not os.path.islink(os.path.join(CC, "SKILL.md")))
for e in ("scripts", "references", "assets", "LICENSE.txt", "THIRD_PARTY_NOTICES.md"):
    path = os.path.join(CC, e)
    check(f"…{e} links into the package", os.path.islink(path) and os.path.realpath(path) == os.path.realpath(os.path.join(PKG, e)), os.readlink(path) if os.path.islink(path) else "missing")

for label, path in (("package SKILL.md", os.path.join(PKG, "SKILL.md")), ("Claude Code SKILL.md", os.path.join(CC, "SKILL.md"))):
    _, fm, body = open(path, encoding="utf-8").read().split("---\n", 2)
    m = yaml.safe_load(fm)
    allowed = {"name", "description", "license", "compatibility", "metadata", "allowed-tools"}
    check(f"{label}: only spec fields in the frontmatter", set(m) <= allowed, str(sorted(set(m) - allowed)))
    check(f"{label}: name = folder name, lowercase/digits/hyphens, <= 64",
          m["name"] == "monoonom" and re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", m["name"]) and len(m["name"]) <= 64)
    d = m["description"].strip()
    check(f"{label}: description 1–1024 chars, no XML tags, third person", 0 < len(d) <= 1024 and not re.search(r"<[^>]+>", d) and not re.match(r"(I|You)\b", d), f"{len(d)} chars")
    check(f"{label}: compatibility <= 500 chars", len(m.get("compatibility", "").strip()) <= 500, f"{len(m.get('compatibility', '').strip())} chars")
    check(f"{label}: metadata values are strings", all(isinstance(v, str) for v in (m.get("metadata") or {}).values()))
    check(f"{label}: body under 500 lines", body.count("\n") < 500, f"{body.count(chr(10))} lines")
    refs = re.findall(r"`(?:\$\{CLAUDE_SKILL_DIR\}/)?((?:references|scripts)/[\w./-]+)`", body)
    base = os.path.dirname(path)
    check(f"{label}: every referenced file exists", all(os.path.exists(os.path.join(base, r)) for r in refs), str(refs))

print(f"{fails} FAILED" if fails else "ALL PASSED")
sys.exit(1 if fails else 0)
