"""Generate SKILL.claude-code.md from the portable SKILL.md (keeps both in sync).

    python3 tools/build_claude_code_skill.py           # write both files
    python3 tools/build_claude_code_skill.py --check   # exit 1 if either is out of date

Differences: Claude Code-only `allowed-tools` with ${CLAUDE_SKILL_DIR}, a live status injection,
and the launch.json path for the Claude Code desktop preview.
"""
import os
import re
import sys

SKILL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".agents", "skills", "monoonom")
src = open(os.path.join(SKILL, "SKILL.md"), encoding="utf-8").read()
_, front, body = src.split("---\n", 2)

front = re.sub(r"compatibility: >\n(?:  .*\n)+", "compatibility: >\n  Claude Code on the user's machine with Python 3.9+; binds a port on 127.0.0.1. Optional\n  Node.js 18+ computes formulas when no preview is open.\n", front)
front += "allowed-tools: Bash(python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py *) Bash(python ${CLAUDE_SKILL_DIR}/scripts/xl.py *)\n"

body = body.replace("# MoNoonOm\n", "# MoNoonOm (Claude Code)\n\nServer status right now (the preview URL with its token comes from `xl.py url`):\n!`python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py status --brief || true`\n", 1)
old = "Run commands as `python scripts/xl.py <command>` using this skill folder's absolute path, **from the"
assert old in body, "SKILL.md run instruction changed; update tools/build_claude_code_skill.py"
body = body.replace(old, "Run commands as `python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py <command>` (written `xl.py` below), **from the")
start = "1. **Start**: `xl.py serve` (reuses a running server). Give the user the URL it prints (see Preview"
assert start in body, "SKILL.md start step changed; update tools/build_claude_code_skill.py"
body = body.replace(start, "1. **Start**: in Claude Code desktop run `xl.py launch-config`, then start the \"monoonom\" preview\n"
                    "   (the app owns the server and stops it with the session). In a terminal session run `xl.py serve`\n"
                    "   and give the user the URL it prints (see Preview")
body = body.replace("`references/", "`${CLAUDE_SKILL_DIR}/references/")
out = "---\n" + front + "---\n" + body

# Where it goes: SKILL.claude-code.md in the package (to install by hand), and the SKILL.md of the
# project's Claude Code skill folder .claude/skills/monoonom/, whose other entries link to the package.
ROOT = os.path.dirname(SKILL)
ROOT = os.path.dirname(os.path.dirname(ROOT))
TARGETS = [os.path.join(SKILL, "SKILL.claude-code.md"), os.path.join(ROOT, ".claude", "skills", "monoonom", "SKILL.md")]

if "--check" in sys.argv:  # tests: fail when SKILL.md changed but the generated files were not rebuilt
    stale = [t for t in TARGETS if not os.path.exists(t) or open(t, encoding="utf-8").read() != out]
    for t in stale:
        print("stale:", os.path.relpath(t, ROOT), "(run python3 tools/build_claude_code_skill.py)")
    sys.exit(1 if stale else 0)
for t in TARGETS:
    with open(t, "w", encoding="utf-8") as f:
        f.write(out)
    print("wrote", os.path.relpath(t, ROOT), len(out.splitlines()), "lines")
