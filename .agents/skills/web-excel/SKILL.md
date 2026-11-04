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
  Requires Python 3.9+ and a local agent session on the user's machine with a shell; binds a
  port on 127.0.0.1 (sandboxes that block loopback binding, e.g. Codex workspace-write, need
  approval or full access). Not usable from the Skills API or cloud sessions, which cannot serve
  the user's localhost.
metadata:
  version: "0.2"
---

# web-excel

A local web spreadsheet that the user sees in the preview pane. All file reads and writes are done
by the bundled Python server (never by the browser), which keeps DRM-protected files working.

Run every command with the Python interpreter from this skill folder:
`python scripts/xl.py <command>` (use the absolute path of this skill directory).

## Start and show the preview

1. `python scripts/xl.py serve` — starts the server for the current workspace, or reuses the
   running one. It prints the URL (default `http://localhost:8765/`).
2. Tell the user to open that URL in the preview pane.
3. `python scripts/xl.py status` shows the URL and open documents at any time.

The server stops by itself after an hour without clients or requests. Stop it explicitly with
`python scripts/xl.py stop` only when the user asks; it refuses while a workbook has unsaved changes.

## When something fails

Read `references/troubleshooting.md` (port conflicts, blocked loopback, DRM-encrypted files).
