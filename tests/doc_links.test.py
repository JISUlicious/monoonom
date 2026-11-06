"""Every relative link and image in the repository's Markdown files points at a file that exists.
    python3 tests/doc_links.test.py
"""
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
files = subprocess.check_output(["git", "ls-files", "*.md"], cwd=ROOT, text=True).split()
LINK = re.compile(r"!?\[[^\]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
broken = []
for f in files:
    text = open(os.path.join(ROOT, f), encoding="utf-8").read()
    text = re.sub(r"```.*?```", "", text, flags=re.S)  # code blocks are not links
    for m in LINK.finditer(text):
        target = m.group(1).split("#")[0]
        if not target or re.match(r"^[a-z]+:", target):  # anchors, http(s):, mailto:
            continue
        if not os.path.exists(os.path.normpath(os.path.join(ROOT, os.path.dirname(f), target))):
            broken.append(f"{f}: {m.group(1)}")
print(f"{len(files)} Markdown files")
for b in broken:
    print("BROKEN", b)
print(f"{len(broken)} FAILED" if broken else "ALL PASSED")
sys.exit(1 if broken else 0)
