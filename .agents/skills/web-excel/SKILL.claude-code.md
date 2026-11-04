---
name: web-excel
description: >
  Opens, inspects, edits, formats, charts, filters and saves Excel workbooks (.xlsx, .xlsm, .xls)
  in the workspace through a local web-excel server, while the user watches and edits the same
  workbook live in the agent's preview pane at a localhost URL. Use when the user asks to read or
  change spreadsheet data, build formulas or charts, or wants to see or edit an Excel file
  together with the agent. Do not use for CSV-only processing, Google Sheets, or one-off
  conversions where no live view is wanted; a plain spreadsheet library is simpler there.
license: MIT
compatibility: >
  Claude Code on the user's machine with Python 3.9+. Binds a port on 127.0.0.1.
metadata:
  version: "0.2"
allowed-tools: Bash(python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py *) Bash(python ${CLAUDE_SKILL_DIR}/scripts/xl.py *)
---

# web-excel (Claude Code)

Current server state:
!`python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py status --brief || true`

All file reads and writes are done by the bundled Python server, never by the browser (DRM-safe).
Commands: `python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py <command>`.

## Start and show the preview

- Claude Code desktop: run `xl.py launch-config`, then start the "web-excel" preview. The app owns the
  server process and stops it with the session.
- Terminal / other hosts: run `xl.py serve` and tell the user the printed URL.

The server stops by itself after an hour without clients or requests. Use `xl.py stop` only when the
user asks; it refuses while a workbook has unsaved changes.

## When something fails

Read `${CLAUDE_SKILL_DIR}/references/troubleshooting.md`.
