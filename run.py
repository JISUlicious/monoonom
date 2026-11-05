#!/usr/bin/env python3
"""Development launcher: runs the skill's server in the foreground.

    python run.py [--port 8765] [--root ./samples]

The skill itself is in .agents/skills/monoonom (entry point: scripts/xl.py).
"""
import os
import sys

XL = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".agents", "skills", "monoonom", "scripts", "xl.py")

if __name__ == "__main__":
    args = sys.argv[1:]
    if "--root" not in args:
        args += ["--root", os.getcwd()]
    os.execv(sys.executable, [sys.executable, XL, "serve", "--foreground", "--idle-timeout", "0", *args])
