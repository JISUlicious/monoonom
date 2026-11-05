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
  Requires Python 3.9+ with openpyxl (pip install -r scripts/requirements.txt; xlrd optional for
  .xls) and a local agent session on the user's machine with a shell; binds a
  port on 127.0.0.1 (sandboxes that block loopback binding, e.g. Codex workspace-write, need
  approval or full access). Optional Node.js 18+ computes formulas when no preview is open.
  Not usable from the Skills API or cloud sessions, which cannot serve the user's localhost.
metadata:
  version: "0.3"
---

# web-excel

A local web spreadsheet shared by the agent and the user. The bundled Python server owns every
open workbook: the agent edits it with `scripts/xl.py`, the user edits it in the preview pane, and
both see each other's changes immediately. All file reads and writes are done by Python, never by
the browser, which keeps DRM-protected files working.

Run commands as `python scripts/xl.py <command>` using this skill folder's absolute path, **from the
user's workspace (project folder) as the current directory — never `cd` into the skill folder**. The
workspace is where the server keeps its state (`.web-excel/`), which files it may open, and what
relative file paths are resolved against. If the shell is elsewhere, `cd` to the workspace first or
pass `--root <workspace>`; xl.py refuses to run with the skill folder as its workspace (exit code 2).

## Preview URL (always with its token)

Give the user **exactly the URL that `xl.py serve` or `xl.py url` prints**, copied verbatim, including
the `#t=…` token at the end. Never write `http://localhost:PORT/` yourself or shorten the URL: without
the token the page loads but no workbook opens. If you no longer have the URL, run `xl.py url`.

## Workflow

1. **Start**: `xl.py serve` (reuses a running server). Give the user the URL it prints (see Preview
   URL above) to open in the preview pane. If it exits with code 6, a Python package is missing: ask
   the user before running the printed `pip install -r scripts/requirements.txt` command, then retry.
2. **Open**: `xl.py open report.xlsx` — the preview shows the file (a tab in every preview window); run
   it again to bring an already open file to the front. If the user is typing in a cell, its tab is
   only marked.
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
   The previous file is backed up to `.web-excel/backups/` first. `.xls` files are read-only: saving
   one writes `NAME.xlsx` next to it, so tell the user the new file name.

Leave the server running when done so the user can keep working; it stops itself after an hour
without clients. Use `xl.py stop` only on request (it refuses while there are unsaved changes).

## Commands

| Task | Command |
|---|---|
| Write values (a row) | `xl.py write B2 100 200 '=B2+C2'` (`--down` for a column) |
| Write a block | `xl.py write A1 --data '[["Name","Qty"],["Apple",3]]'` or `--tsv file.tsv` |
| Fill a range with one formula | `xl.py fill 'D2:D100' '=B2*C2'` (relative refs shift per row) |
| Clear | `xl.py clear A1:C9 --what contents\|formats\|comments\|all` |
| Format | `xl.py format A1:F1 b=true bg=#DDEBF7 ha=center nf='#,##0' --border outline` |
| Column width / row height | `xl.py width A:C 14` (characters), `xl.py height 1 24` (points) |
| Merge / freeze / filter | `xl.py merge A1:F1`, `xl.py freeze B2`, `xl.py filter A1:F50` (`off`) |
| Insert / delete rows, columns | `xl.py rows insert 5 --count 2`, `rows delete 5:7`, `cols insert C`, `cols delete C:D` |
| Insert / delete cells (shift) | `xl.py cells insert B2:C3 --shift down`, `cells delete B2 --shift left` |
| Sort / remove duplicates | `xl.py sort A1:F50 --by C:desc --by A`, `xl.py dedupe A1:F50 --columns A,C` |
| Filter conditions | `xl.py filter col B --where '>150'`, `filter col A --values '서울\|부산'`, `filter clear`, `filter reapply` |
| Copy / paste special / move | `xl.py copy A1:C5 E1 [--what values\|formats\|formulas] [--transpose]`, `xl.py move A1:C5 E1` |
| Series (fill handle) | `xl.py series A2:A3 A2:A50` (1,2→…, 월→화, 항목1→항목2, dates +1) |
| Find / replace | `xl.py find '*서울*'`, `xl.py replace 구 신 --all [--in formulas]` |
| Notes, links, autofit | `xl.py comment B2 '확인'`, `xl.py link A1 https://… --text 사이트`, `xl.py autofit A:F` |
| Format copy, decimals | `xl.py format-copy A1:F1 A20:F20`, `xl.py decimals B2:B9 +1` |
| Sheets | `xl.py sheet add 요약`, `rename Old New`, `rm`, `copy`, `move Name 1`, `hide`/`unhide`, `color Name #FF0000`, `gridlines Name off` |
| Charts | `xl.py chart add 'A1:C13' --type line --title 'Monthly' --at H2`, `chart set ID --title … --to-type bar --data A1:D13`, `chart ls`, `chart rm ID` |
| Preview URL (with token) | `xl.py url` |
| Show the user something | `xl.py goto 'Sheet2!B10'` (moves the preview), `xl.py view --zoom 150 --formulas` |
| Follow mode | The preview follows your edits to their sheet and cells by default (`xl.py view --follow on\|off`); when off, your cells are only outlined |
| History / undo | `xl.py changes --since 0`, `xl.py undo BATCH` |

Tables do not have to start at A1. Give filter/sort/chart/dedupe a single cell or whole columns and
they use the table there (the contiguous block, without a 합계/Total row under it); an explicit
multi-cell range is used as given. `describe` lists each table with its header and totals rows.

Row/column inserts, deletes and sorts move formulas, chart ranges, merges and the filter range just
like Excel, everywhere in the workbook. If the user edits cells at the same moment, their edits are
moved to the shifted position automatically.

Values are interpreted like typing into Excel: `=` starts a formula, `2026-03-15` becomes a date,
`15%` a percentage, `1,234` a number, `'007` stays text. Pass `--raw` to keep strings as text.
`read` shows dates as `YYYY-MM-DD`; `--what text` shows exactly what the grid displays (number
formats applied, needed for filter `--values`); `--what raw|formulas|all` shows serials, formulas or styles.

Format keys: `b i u st` (bold, italic, underline, strike), `fs` size, `ff` font, `fc` font color,
`bg` fill, `ha` left/center/right, `va` top/center/bottom, `wrap`, `ind`, `nf` number format.
Chart types: bar, barH, barStacked, line, area, pie, doughnut, scatter, radar.

When several workbooks are open, add `--doc NAME_OR_ID`. Add `--json` for machine-readable output.
Full option list: `references/cli.md`. Batch edit format for `apply`: `references/edits.md`.

## Formula results

Formula values come from the open preview (the browser calculates). If no preview is open, Node.js
computes them; without Node, `read` marks values as stale and they may be the file's last saved
results — then ask the user to open the preview.

Functions the engine cannot calculate (LAMBDA, VBA user functions, newer Excel functions) are
reported when the file opens and in `describe`. Those cells show the file's last result and are
saved with their original formula unchanged as long as the formula is not edited; tell the user
before changing such a formula, since its result can then only be calculated by Excel.

## Exit codes and problems

0 ok, 2 usage (e.g. ambiguous `--doc`), 3 server not running, 4 conflict (file changed on disk,
name exists, unsaved changes), 5 file/DRM/sheet not found, 6 Python package missing. See `references/troubleshooting.md`.
