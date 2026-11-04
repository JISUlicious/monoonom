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
  Claude Code on the user's machine with Python 3.9+; binds a port on 127.0.0.1. Optional
  Node.js 18+ computes formulas when no preview is open.
metadata:
  version: "0.3"
allowed-tools: Bash(python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py *) Bash(python ${CLAUDE_SKILL_DIR}/scripts/xl.py *)
---

# web-excel (Claude Code)

Server status right now:
!`python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py status --brief || true`

A local web spreadsheet shared by the agent and the user. The bundled Python server owns every
open workbook: the agent edits it with `scripts/xl.py`, the user edits it in the preview pane, and
both see each other's changes immediately. All file reads and writes are done by Python, never by
the browser, which keeps DRM-protected files working.

Run commands as `python3 ${CLAUDE_SKILL_DIR}/scripts/xl.py <command>` (written `xl.py` below). Relative
file paths are resolved against the workspace (the current project folder).

## Workflow

1. **Start**: in Claude Code desktop run `xl.py launch-config`, then start the "web-excel" preview
   (the app owns the server and stops it with the session). In a terminal session run `xl.py serve`
   and tell the user the printed URL (default `http://localhost:8765/`).
2. **Open**: `xl.py open report.xlsx` — the file also appears as a tab in the preview.
   `xl.py ls` lists Excel files in the workspace.
3. **Understand before editing**: `xl.py describe` (sheets, used range, header row, column types,
   charts), then `xl.py read 'Sheet1!A1:F20'` for just the range you need. Avoid reading whole
   sheets; `read` stops at 200 rows and prints where to continue.
4. **Edit**: use the specific commands below, or `xl.py apply edits.json` for several changes at once
   (one undoable batch). Every edit prints a batch id; `xl.py undo BATCH` reverts it.
5. **Check the user's side**: the user may be editing at the same time. Before changing an area the
   user is working on, run `xl.py changes --since VERSION --actor user` and `xl.py selection`
   ("this part" usually means the user's current selection).
6. **Save only when the user asks** (or agreed): `xl.py save` or `xl.py save --as new.xlsx`.
   The previous file is backed up to `.web-excel/backups/` first.

Leave the server running when done so the user can keep working; it stops itself after an hour
without clients. Use `xl.py stop` only on request (it refuses while there are unsaved changes).

## Commands

| Task | Command |
|---|---|
| Write values (a row) | `xl.py write B2 100 200 '=B2+C2'` (`--down` for a column) |
| Write a block | `xl.py write A1 --data '[["Name","Qty"],["Apple",3]]'` or `--tsv file.tsv` |
| Fill a range with one formula | `xl.py fill 'D2:D100' '=B2*C2'` (relative refs shift per row) |
| Clear | `xl.py clear A1:C9 --what contents\|formats\|all` |
| Format | `xl.py format A1:F1 b=true bg=#DDEBF7 ha=center nf='#,##0' --border outline` |
| Column width / row height | `xl.py width A:C 14` (characters), `xl.py height 1 24` (points) |
| Merge / freeze / filter | `xl.py merge A1:F1`, `xl.py freeze B2`, `xl.py filter A1:F50` (`off`) |
| Insert / delete rows, columns | `xl.py rows insert 5 --count 2`, `rows delete 5:7`, `cols insert C`, `cols delete C:D` |
| Sort | `xl.py sort A1:F50 --by C:desc --by A` (header row detected; `--header`/`--no-header`) |
| Sheets | `xl.py sheet add 요약`, `sheet rename Old New` (formulas follow), `sheet rm Name` |
| Charts | `xl.py chart add 'A1:C13' --type line --title 'Monthly' --at H2`, `chart ls`, `chart rm ID` |
| History / undo | `xl.py changes --since 0`, `xl.py undo BATCH` |

Row/column inserts, deletes and sorts move formulas, chart ranges, merges and the filter range just
like Excel, everywhere in the workbook. If the user edits cells at the same moment, their edits are
moved to the shifted position automatically.

Values are interpreted like typing into Excel: `=` starts a formula, `2026-03-15` becomes a date,
`15%` a percentage, `1,234` a number, `'007` stays text. Pass `--raw` to keep strings as text.
`read` shows dates as `YYYY-MM-DD`; `--what raw|formulas|all` shows serials, formulas or styles.

Format keys: `b i u st` (bold, italic, underline, strike), `fs` size, `ff` font, `fc` font color,
`bg` fill, `ha` left/center/right, `va` top/center/bottom, `wrap`, `ind`, `nf` number format.
Chart types: bar, barH, barStacked, line, area, pie, doughnut, scatter, radar.

When several workbooks are open, add `--doc NAME_OR_ID`. Add `--json` for machine-readable output.
Full option list: `${CLAUDE_SKILL_DIR}/references/cli.md`. Batch edit format for `apply`: `${CLAUDE_SKILL_DIR}/references/edits.md`.

## Formula results

Formula values come from the open preview (the browser calculates). If no preview is open, Node.js
computes them; without Node, `read` marks values as stale and they may be the file's last saved
results — then ask the user to open the preview.

## Exit codes and problems

0 ok, 2 usage (e.g. ambiguous `--doc`), 3 server not running, 4 conflict (file changed on disk,
name exists, unsaved changes), 5 file/DRM/sheet not found. See `${CLAUDE_SKILL_DIR}/references/troubleshooting.md`.
