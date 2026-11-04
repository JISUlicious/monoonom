# xl.py command reference

## Contents
- Server: serve, status, stop, launch-config
- Documents: ls, open, new, docs, close, save
- Reading: describe, read, changes, selection
- Editing: write, fill, clear, format, merge, unmerge, width, height, hide, unhide, freeze, filter, rows, cols, sort, sheet, chart, apply, undo
- Common options and exit codes

Every command accepts `--root DIR` (workspace; default `$WEBEXCEL_ROOT`, `$CLAUDE_PROJECT_DIR` or the
current folder). Document commands accept `--doc ID|NAME|PATH` (optional when one workbook is open)
and most accept `--json`.

## Server

| Command | Notes |
|---|---|
| `serve [--port P] [--idle-timeout S] [--no-sandbox] [--allow-host H]` | Starts a detached server for the workspace or reuses the running one. Picks the first free port from 8765. Stops after S seconds (default 3600) with no clients/requests; dirty workbooks are written to `.web-excel/recovery/` first. |
| `serve --foreground [--exit-with-parent]` | Runs in the current process (used by `.claude/launch.json`); exits when the parent exits. |
| `status [--brief]` | URL, pid, open workbooks. Exit 3 if not running (`--brief` always exits 0). |
| `stop [--force]` | Refuses (exit 4) while a workbook has unsaved changes unless `--force`. |
| `launch-config [--port P]` | Adds/updates the `web-excel` entry in `.claude/launch.json` (Claude Code desktop preview owns the process). |

## Documents

| Command | Notes |
|---|---|
| `ls [DIR]` | Excel files and folders. |
| `open PATH` | Opens (or reuses) a workbook; it appears as a tab in every preview window. |
| `new [--name N]` | New blank workbook (needs `save --as` later). |
| `docs` | Open workbooks: id, version, unsaved flag, number of preview windows. |
| `save [--as PATH] [--format xlsx\|xlsm\|xls] [--overwrite] [--force] [--no-backup]` | Python writes the file. `--overwrite`: replace another existing file. `--force`: overwrite although the file changed on disk since it was opened. Backup goes to `.web-excel/backups/` unless `--no-backup` (for DRM policies that forbid copies). |
| `close [--force]` | Closes for everyone; refuses with unsaved changes unless `--force`. |

xlsx saving keeps parts the editor does not show (conditional formatting, data validation, images,
VBA in .xlsm). `.xls` cannot store charts or autofilter (warnings are printed).

## Reading

| Command | Notes |
|---|---|
| `describe` | Per sheet: used range, cell/formula counts, header row, column types, freeze, filter, merges, charts (id, type, anchor, data refs), defined names, calc status. |
| `read RANGE [--what values\|raw\|formulas\|all] [--max N]` | `values`: dates as ISO text; `raw`: serial numbers; `formulas`: formula text or value; `all`: value + formula + style. Full-column ranges are cut to the used area; output stops at `--max` rows (default 200) and prints the remaining range. |
| `changes [--since V] [--actor user\|agent]` | Batches after version V: who, label, batch id, changed ranges. |
| `selection` | The user's current selection in each preview window (most recent first). |

## Editing

All edits are one batch per command, applied live and undoable with `undo BATCH`.

| Command | Notes |
|---|---|
| `write RANGE V... [--down] [--data JSON] [--tsv FILE\|-] [--raw]` | Values parse like Excel input. A single value with a multi-cell range fills it. |
| `fill RANGE VALUE` | Same value/formula everywhere; relative references shift (`=B2*C2` → `=B3*C3`). |
| `clear RANGE [--what contents\|formats\|all]` | `contents` keeps formatting. |
| `format RANGE k=v ... [--unset k] [--border KIND[:STYLE[:#COLOR]]]` | Border kinds: all, outline, inside, top, bottom, left, right, none. Styles: thin, medium, thick, dashed, dotted, double, hair. |
| `merge RANGE` / `unmerge RANGE` | Merge keeps the top-left value and centers it. |
| `width COLS CHARS` / `height ROWS POINTS` | e.g. `width B:D 12`, `height 1:3 20`. |
| `hide RANGE` / `unhide RANGE` | Columns (`B:C`) or rows (`3:5`). |
| `freeze CELL\|none [--sheet S]` | `B2` freezes row 1 and column A. |
| `filter RANGE\|off [--sheet S]` | Turns the autofilter on for a header + data range. |
| `rows insert AT [--count N] [--sheet S]` / `rows delete 5:7` | Inserts before row AT. Formulas in every sheet, chart ranges, merges, the filter range and hidden rows follow (same rules as Excel and the browser). |
| `cols insert C [--count N]` / `cols delete C:D` | Same for columns. |
| `sort RANGE --by COL[:asc\|desc] [--by ...] [--header\|--no-header]` | Header row is detected when omitted; rows hidden by the filter stay put; blanks sort last; formulas move with their row. |
| `sheet add NAME` / `sheet rename OLD NEW` / `sheet rm NAME` | Rename rewrites formulas, chart ranges and defined names. |
| `chart add DATA [--type T] [--title] [--at CELL] [--sheet S] [--legend r\|b\|t\|l] [--width PX] [--height PX]` | First row = series names, first text/date column = categories (like Excel). |
| `chart ls` / `chart rm ID [--sheet S]` | |
| `apply FILE\|- [--label TEXT]` | Several edits as one batch; see `edits.md`. |
| `undo BATCH [--force]` | Refuses if those cells were changed again later, unless `--force`. |

## Exit codes

0 ok · 1 other error · 2 usage · 3 server not running · 4 conflict (409) · 5 file, DRM, path or sheet error.
