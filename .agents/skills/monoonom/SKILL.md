---
name: monoonom
description: >
  Opens, inspects, edits, formats, charts, filters and saves Excel workbooks (.xlsx, .xlsm, .xls;
  .csv and .tsv open as one sheet) in the workspace through a local monoonom server, while the
  user watches and edits the same workbook live in the agent's preview pane at a localhost URL.
  Use when the user asks to read or change spreadsheet data, build formulas or charts, or wants
  to see or edit an Excel file together with the agent. Do not use for batch CSV processing,
  Google Sheets, or one-off conversions where no live view is wanted; a plain spreadsheet
  library is simpler there.
license: MIT
compatibility: >
  Requires Python 3.9+ with openpyxl and Pillow (pip install -r scripts/requirements.txt; xlrd
  optional for .xls) and a local agent session on the user's machine with a shell; binds a
  port on 127.0.0.1 (sandboxes that block loopback binding, e.g. Codex workspace-write, need
  approval or full access). Optional Node.js 18+ computes formulas when no preview is open.
  Not usable from the Skills API or cloud sessions, which cannot serve the user's localhost.
metadata:
  version: "0.3"
---

# MoNooNoM

A local web spreadsheet shared by the agent and the user. The bundled Python server owns every
open workbook: the agent edits it with `scripts/xl.py`, the user edits it in the preview pane, and
both see each other's changes immediately. All file reads and writes are done by Python, never by
the browser, which keeps DRM-protected files working.

Run commands as `python scripts/xl.py <command>` using this skill folder's absolute path, **from the
user's workspace (project folder) as the current directory — never `cd` into the skill folder**. The
workspace is where the server keeps its state (`.monoonom/`), which files it may open, and what
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
   The previous file is backed up to `.monoonom/backups/` first. `.xls` files are read-only: saving
   one writes `NAME.xlsx` next to it, so tell the user the new file name. A `.csv`/`.tsv` file
   opens as one sheet and a plain `save` likewise writes `NAME.xlsx` (the `.csv` stays unchanged);
   `save --as NAME.csv` writes only the active sheet's values, so use it only when asked.

Leave the server running when done so the user can keep working; it stops itself after an hour
without clients. Use `xl.py stop` only on request (it refuses while there are unsaved changes).

## Analysis and data processing: in the workbook, with formulas

When the user asks to calculate, summarize, aggregate, look up, clean or reshape data, do it **in the
workbook with Excel formulas** so the result stays live, can be checked in the preview and still
works when the file is opened in Excel. Do not compute it in Python (pandas, scripts) and paste the
values in.

- Aggregates and summaries: `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`/`MINIFS`, `SUMPRODUCT`;
  a summary table = a column of unique keys plus `SUMIFS` next to it (`fill` copies the formula down).
- Lookups and joins: `XLOOKUP` (or `INDEX`/`MATCH`); ranking and top-N: `RANK.EQ`, `LARGE`/`SMALL`
  with `INDEX`/`MATCH` for the matching labels.
- No dynamic-array functions (`UNIQUE`, `SORT`, `SORTBY`, `FILTER`, `SEQUENCE`): their results do not
  spill here (only the top-left value shows) and are saved as single-cell formulas. For a list of
  unique keys, `read` the column and write the keys as values, then `SUMIFS`/`COUNTIFS` next to them.
- Cleaning and reshaping: `TRIM`, `CLEAN`, `UPPER`/`PROPER`, `SUBSTITUTE`, `LEFT`/`MID`/`RIGHT`,
  `TEXTBEFORE`/`TEXTAFTER`, `VALUE`/`DATEVALUE`, `IF`/`IFS`/`IFERROR`; add helper columns rather
  than overwriting the source data.
- Write one formula with relative references and `fill` it over the range instead of writing each
  cell; then `read` a few result cells to check the values.
- Put constants the user may want to change (thresholds, bin edges, rates, targets) in labelled
  input cells and refer to them (`$B$1`) instead of typing them into formulas.
- Before finishing, `read --what formulas` the result cells: anything derived from other cells must
  be a formula, not a typed number.
- Use only functions this engine calculates: `describe` reports unsupported ones, and a new formula
  with one shows `#NAME?` here (Excel still calculates it). Prefer a supported equivalent.
- Python is the exception: only when Excel functions cannot express it (e.g. fetching external data,
  heavy statistics, more rows than a sheet allows) or the user asks for static values. Then say so,
  and write the inputs and results into clearly labelled cells.

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
| Filter conditions | `xl.py filter col B --where '>150'`, `filter col A --values 'Seoul\|Busan'`, `filter clear`, `filter reapply` |
| Copy / paste special / move | `xl.py copy A1:C5 E1 [--what values\|formats\|formulas] [--transpose]`, `xl.py move A1:C5 E1` |
| Series (fill handle) | `xl.py series A2:A3 A2:A50` (1,2→…, Mon→Tue, Item1→Item2, dates +1) |
| Find / replace | `xl.py find '*Seoul*'`, `xl.py replace old new --all [--in formulas]` |
| Notes, links, autofit | `xl.py comment B2 'Check'`, `xl.py link A1 https://… --text Site`, `xl.py autofit A:F` |
| Format copy, decimals | `xl.py format-copy A1:F1 A20:F20`, `xl.py decimals B2:B9 +1` |
| Sheets | `xl.py sheet add Summary`, `rename Old New`, `rm`, `copy`, `move Name 1`, `hide`/`unhide`, `color Name #FF0000`, `gridlines Name off` |
| Pictures | `xl.py image add logo.png --at H2 [--width 200]` (a picture file in the workspace; the user can also paste pictures in the preview), `image ls`, `image set ID --at B20 --width 300`, `image rm ID` |
| Charts | `xl.py chart add 'A1:C13' --type line --title 'Monthly' --at H2`, `chart set ID --title … --to-type bar --data A1:D13` (`--switch` swaps rows/columns), `chart ls`, `chart rm ID`. Chart ids change whenever the file is reopened or the server restarts: run `chart ls` before `chart set/rm` instead of reusing an older id |
| Preview URL (with token) | `xl.py url` |
| Show the user something | `xl.py goto 'Sheet2!B10'` (moves the preview), `xl.py view --zoom 150 --formulas` |
| Follow mode | The preview follows your edits to their sheet and cells by default (`xl.py view --follow on\|off`); when off, your cells are only outlined |
| History / undo | `xl.py changes --since 0`, `xl.py undo BATCH` |

Tables do not have to start at A1. Give filter/sort/chart/dedupe a single cell or whole columns and
they use the table there (the contiguous block, without a Total row under it); an explicit
multi-cell range is used as given. `describe` lists each table with its header and totals rows.

Pivot tables in the file are kept and listed by `describe` / `xl.py pivot ls` (range, source,
fields). Their output cells cannot be edited, sorted or cut through by row/column inserts (the
command fails with exit code 4 and says why): change the pivot's source data instead, then run
`xl.py pivot refresh [NAME]` to recompute its values here (one undoable batch; it prints notes,
e.g. new items that only Excel can add rows for). Saving makes Excel refresh the pivot when it
opens the file either way. Inserting rows/columns above or left of a pivot moves it.
To make one: `xl.py pivot add 'Data!A1' --rows Region --cols Product --values Sales:sum
[--filters Quarter] --at 'Pivot!A3'` (fields are the source's header names; a missing sheet is
created; several `--rows`/`--values` separated by commas; with `--cols` only one value field; `--lang ko` for Korean labels when the user works
in Korean Excel). It is written in tabular form and saved as a real pivot table that Excel rebuilds when it opens the
file. Prefer a pivot when the user asks for one; for a summary they will edit, use formulas.
Pivot charts (`chart ls` says "pivot chart of 'NAME'") stay linked to their pivot when edited; their
data range, row/column switch and scatter type cannot be changed (change the pivot). To make one:
`xl.py chart add --pivot NAME [--type line] [--at 'S!H2']` (pivots with one row field).

Row/column inserts, deletes and sorts move formulas, chart ranges, merges and the filter range just
like Excel, everywhere in the workbook. If the user edits cells at the same moment, their edits are
moved to the shifted position automatically.

Values are interpreted like typing into Excel: `=` starts a formula, `2026-03-15` becomes a date,
`15%` a percentage, `1,234` a number, `'007` stays text. Pass `--raw` to keep strings as text.
`read` shows dates as `YYYY-MM-DD`; `--what text` shows exactly what the grid displays (number
formats applied, needed for filter `--values`); `--what raw|formulas|all` shows serials, formulas or styles.

Format keys: `b i u st` (bold, italic, underline, strike), `fs` size (1-409), `ff` font (`xl.py fonts` lists the installed ones), `fc` font color,
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
