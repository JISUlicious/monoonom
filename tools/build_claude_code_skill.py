"""Generate SKILL.claude-code.md from the portable SKILL.md (keeps both in sync).

    python3 tools/build_claude_code_skill.py

Differences: Claude Code-only `allowed-tools` with ${CLAUDE_SKILL_DIR}, a live status injection,
and the launch.json path for the Claude Code desktop preview.
"""
import os
import re

SKILL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".agents", "skills", "web-excel")
src = open(os.path.join(SKILL, "SKILL.md"), encoding="utf-8").read()
_, front, body = src.split("---\n", 2)

front = re.sub(r"compatibility: >\n(?:  .*\n)+", "compatibility: >\n  Claude Code on the user's machine with Python 3.9+; binds a port on 127.0.0.1. Optional\n  Node.js 18+ computes formulas when no preview is open.\n", front)
front += "allowed-tools: Bash(python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py *) Bash(python ${CLAUDE_SKILL_DIR}/scripts/xl.py *)\n"

body = body.replace("# web-excel\n", "# web-excel (Claude Code)\n\nServer status right now:\n!`python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py status --brief || true`\n", 1)
old = "Run commands as `python scripts/xl.py <command>` using this skill folder's absolute path, **from the"
assert old in body, "SKILL.md run instruction changed; update tools/build_claude_code_skill.py"
body = body.replace(old, "Run commands as `python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py <command>` (written `xl.py` below), **from the")
start = "1. **Start**: `xl.py serve` (reuses a running server). Give the user the printed URL **including its"
assert start in body, "SKILL.md start step changed; update tools/build_claude_code_skill.py"
body = body.replace(start, "1. **Start**: in Claude Code desktop run `xl.py launch-config`, then start the \"web-excel\" preview\n"
                    "   (the app owns the server and stops it with the session). In a terminal session run `xl.py serve`\n"
                    "   and give the user the printed URL **including its")
body = body.replace("`references/", "`${CLAUDE_SKILL_DIR}/references/")
out = "---\n" + front + "---\n" + body
open(os.path.join(SKILL, "SKILL.claude-code.md"), "w", encoding="utf-8").write(out)
print("wrote SKILL.claude-code.md", len(out.splitlines()), "lines")
