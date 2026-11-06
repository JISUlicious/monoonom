# xl.py command reference

## Contents
- Server: serve, url, status, stop, launch-config
- Documents: ls, open, new, docs, close, save
- Reading: describe, read, changes, selection
- Editing: write, fill, series, clear, format, format-copy, decimals, merge, unmerge, width, height, autofit, hide, unhide, freeze, filter, rows, cols, cells, sort, dedupe, copy, move, find, replace, comment, link, sheet, chart, apply, undo
- View: goto, view
- Common options and exit codes

Run xl.py with the user's workspace as the current folder (call it by the skill's absolute path; do not
`cd` into the skill folder). Every command accepts `--root DIR` (workspace; default `$MONOONOM_ROOT`,
`$CLAUDE_PROJECT_DIR` or the current folder); a workspace inside the skill folder is refused (exit 2). Document commands accept `--doc ID|NAME|PATH` (optional when one workbook is open)
and most accept `--json`.

## Server

| Command | Notes |
|---|---|
| `serve [--port P] [--idle-timeout S] [--no-sandbox] [--allow-host H]` | Starts a detached server for the workspace or reuses the running one. Picks the first free port from 8765. Prints the page URL with the workspace token (`http://localhost:PORT/#t=TOKEN`, also in `status`); the token stays the same across restarts (`.monoonom/token`); `serve --new-token` (server stopped) issues a new one, after which old preview URLs stop working. `status --brief` omits the token. Stops after S seconds (default 3600) with no clients/requests; dirty workbooks are written to `.monoonom/recovery/` first. |
| `serve --foreground [--exit-with-parent]` | Runs in the current process (used by `.claude/launch.json`); exits when the parent exits. |
| `url [--bare]` | The preview URL with its token (`http://localhost:PORT/#t=TOKEN`), to give the user verbatim. `--bare`: the URL line only. Exit 3 if not running. |
| `status [--brief]` | URL, pid, open workbooks. Exit 3 if not running (`--brief` always exits 0). |
| `stop [--force]` | Refuses (exit 4) while a workbook has unsaved changes unless `--force`. |
| `launch-config [--port P]` | Adds/updates the `monoonom` entry in `.claude/launch.json` (Claude Code desktop preview owns the process). |

## Documents

| Command | Notes |
|---|---|
| `ls [DIR]` | Excel files and folders. |
| `open PATH` | Opens (or reuses) a workbook and shows it in every preview window (an already open one is brought to the front; a window where the user is typing only marks the tab). |
| `new [--name N] [--lang en\|ko]` | New blank workbook (needs `save --as` later). Default name and body font as Excel in that language: Book1.xlsx / Calibri, or the Korean ones with `--lang ko`. |
| `docs` | Open workbooks: id, version, unsaved flag, number of preview windows. |
| `save [--as PATH] [--format xlsx\|xlsm] [--overwrite] [--force] [--no-backup]` | Python writes the file. `--overwrite`: replace another existing file. `--force`: overwrite although the file changed on disk since it was opened. Backup goes to `.monoonom/backups/` unless `--no-backup` (for DRM policies that forbid copies). |
| `close [--force]` | Closes for everyone; refuses with unsaved changes unless `--force`. |

xlsx saving keeps parts the editor does not show (conditional formatting, data validation, images,
VBA in .xlsm). `.xls` files are read-only: `save` writes `NAME.xlsx` next to the `.xls` (which stays
unchanged) and the document then points to the `.xlsx`; `--as x.xls` is also saved as `x.xlsx`.

## Reading

| Command | Notes |
|---|---|
| `describe` | Per sheet: used range, cell/formula counts, header row, column types, freeze, filter, merges, charts (id, type, anchor, data refs), defined names, calc status. |
| `read RANGE [--what values\|raw\|formulas\|all] [--max N]` | `values`: dates as ISO text; `raw`: serial numbers; `formulas`: formula text or value; `all`: value + formula + style. Full-column ranges are cut to the used area; output stops at `--max` rows (default 200) and prints the remaining range. |
| `changes [--since V] [--actor user\|agent]` | Batches after version V: who, label, batch id, changed ranges. |
| `selection` | The user's current selection in each preview window (most recent first). |
| `fonts [--search TEXT]` | Fonts installed on this computer (read from the font folders; no server needed), with their Korean and English names; any of them works as `ff=` in `format`. |
| `find TEXT [--in values\|formulas] [--match-case] [--whole] [--sheet S] [--max N]` | `*` and `?` are wildcards. |
| `read RANGE --what text` | Exactly the text shown in the grid (number formats applied by the same formatter as the browser; approximate without Node.js). |

## Editing

All edits are one batch per command, applied live and undoable with `undo BATCH`.

| Command | Notes |
|---|---|
| `write RANGE V... [--down] [--data JSON] [--tsv FILE\|-] [--raw]` | Values parse like Excel input. A single value with a multi-cell range fills it. |
| `fill RANGE VALUE` | Same value/formula everywhere; relative references shift (`=B2*C2` → `=B3*C3`). |
| `clear RANGE [--what contents\|formats\|all]` | `contents` keeps formatting. |
| `format RANGE k=v ... [--unset k] [--border KIND[:STYLE[:#COLOR]]]` | Border kinds: all, outline, inside, top, bottom, left, right, none. Styles: thin, medium, thick, dashed, dotted, double, hair. |
| `merge RANGE [--across] [--no-center]` / `unmerge RANGE` | Merge keeps the top-left value and centers it; `--across` merges each row separately. |
| `series SOURCE TARGET` | Fill handle: numbers continue the linear trend (one number copies; one date adds a day), lists (weekdays, months, quarters, in English and Korean), trailing numbers (Item1→Item2), formulas shift. TARGET extends SOURCE in one direction. |
| `copy FROM TO [--what all\|values\|formats\|formulas] [--transpose]` | Paste special. TO is the top-left cell; formulas shift relative references. |
| `move FROM TO` | Cut and paste: cells keep their formulas unchanged. |
| `cells insert RANGE --shift right\|down` / `cells delete RANGE --shift left\|up` | Shifts neighbouring cells (other formulas are not rewritten, as in the UI). |
| `dedupe RANGE [--columns A,C] [--header\|--no-header]` | Remove duplicate rows (case-insensitive). |
| `replace FIND WITH [--all] [--in values\|formulas] [--match-case] [--whole]` | Constants by default; `--in formulas` edits formula text. |
| `comment CELL [TEXT]` | Note; without TEXT deletes it. |
| `link CELL [URL] [--text T]` | Hyperlink (blue underline unless styled); without URL removes it. |
| `autofit COLS\|ROWS` | e.g. `autofit A:F`, `autofit 1:20` (estimates text width; CJK counts double). |
| `format-copy FROM TO` | Format painter; the source pattern is tiled over TO. |
| `decimals RANGE ±N` | Increase/decrease decimal places. |
| `width COLS CHARS` / `height ROWS POINTS` | e.g. `width B:D 12`, `height 1:3 20`. |
| `hide RANGE` / `unhide RANGE` | Columns (`B:C`) or rows (`3:5`). |
| `freeze CELL\|none [--sheet S]` | `B2` freezes row 1 and column A. |
| `filter RANGE\|off [--sheet S]` | Turns the autofilter on for a header + data range. |
| `filter col C --where COND [--where COND] [--any]` | Conditions: `>150`, `<=10`, `<>0`, `=Seoul`, `contains Se`, `notContains x`, `beginsWith A`, `endsWith ul`; AND by default. |
| `filter col C --values 'a\|b' [--blank]` | Show only these displayed values (see `read --what text`). |
| `filter clear` / `filter reapply` | Remove all conditions / re-evaluate and include rows added below. |
| `rows insert AT [--count N] [--sheet S]` / `rows delete 5:7` | Inserts before row AT. Formulas in every sheet, chart ranges, merges, the filter range and hidden rows follow (same rules as Excel and the browser). |
| `cols insert C [--count N]` / `cols delete C:D` | Same for columns. |
| `sort RANGE --by COL[:asc\|desc] [--by ...] [--header\|--no-header]` | Header row is detected when omitted; rows hidden by the filter stay put; blanks sort last; formulas move with their row. |
| `sheet add NAME` / `sheet rename OLD NEW` / `sheet rm NAME` | Rename rewrites formulas, chart ranges and defined names. |
| `sheet copy NAME [NEW]` / `sheet move NAME POS\|BEFORE_SHEET` | |
| `sheet hide NAME` / `sheet unhide NAME` / `sheet color NAME #RRGGBB\|none` / `sheet gridlines NAME on\|off` | |
| `chart add DATA [--type T] [--title] [--at CELL] [--sheet S] [--legend r\|b\|t\|l] [--width PX] [--height PX]` | First row = series names, first text/date column = categories (like Excel). The chart goes on the data's sheet, or on the sheet of `--at 'Sheet!E2'` / `--sheet`. |
| `chart set ID [--to-type T] [--title] [--data RANGE] [--switch] [--legend r\|b\|t\|l] [--at CELL] [--width] [--height] [--x-title] [--y-title]` | Edit an existing chart (a pivot chart refuses `--data`, `--switch` and scatter: its data comes from its pivot table) (ids from `chart ls`; found on any sheet, `--sheet` only narrows the search). `--at 'Other!K2'` moves it to another sheet. `--switch`: switch rows/columns (series ↔ categories, as Excel's Switch Row/Column; series must be one block of the same size; not for scatter). `--vary-colors on\|off`: one-series charts with a color per category (the legend then lists the categories) or one color. |
| `pivot add SOURCE --rows F[,F] --values F:fn[,F:fn] [--cols F] [--filters F] [--at 'S!A3'] [--title NAME] [--lang en\|ko]` | Create a pivot table from SOURCE (a range or one cell of the table; first row = field names). Tabular form, subtotals after each outer row group, grand totals, report filters above it; `--cols` with one value field only. The target area must be empty; a missing sheet is created. Saved as a real pivot table (definition + cache, refreshOnLoad). `--lang ko` writes the labels as Korean Excel does (Grand Total, Sum of … in Korean); default English. |
| `pivot ls` / `pivot refresh [NAME] [--sheet S]` | Pivot tables: range, source, row/column/value/filter fields. `refresh` recomputes the data cells from the source in the file's layout (sum, count, average, max, min, product, countNums, stdDev(p), var(p); filters, subtotals, grand totals). Without Excel's layout records (LibreOffice files) only one row and one column field. New source items are reported, not added: Excel adds them when it opens the saved file. Prints the pivot charts linked to each refreshed pivot (they redraw from its cells). |
| `chart add --pivot NAME [--type T] [--title] [--at CELL] [--sheet S]` | A pivot chart of pivot table NAME, as Excel makes one: one series per data column (column items, or value fields), the row items as categories, grand totals and the Total Result row left out; linked by `c:pivotSource` (named after the saved file). Pivots with one row field only; not scatter. `--sheet`: the pivot's sheet when two pivots share a name. |
| `chart ls` / `chart rm ID [--sheet S]` | `ls` lists every sheet's charts (a pivot chart says "pivot chart of 'NAME'") with one line per series: name cell, X (categories, or x values for scatter) and Y (values) ranges; `rm` finds the id on any sheet. Ids are not stored in the file: they change when the workbook is reopened or the server restarts, so run `chart ls` again before `set`/`rm`. |
| `goto RANGE [--sheet S]` | Moves the user's preview to the sheet/range (does not change the workbook). |
| `view [--zoom PCT] [--formulas\|--no-formulas] [--follow on\|off]` | Preview zoom and formula display. `--follow`: whether open previews follow your edits (switch to the workbook/sheet and scroll to the cells; default on, the user can toggle it in the status bar). Off: the user works independently and your cells are outlined. |
| `apply FILE\|- [--label TEXT]` | Several edits as one batch; see `edits.md`. |
| `undo BATCH [--force]` | Refuses if those cells were changed again later, unless `--force`. |

## Exit codes

0 ok · 1 other error · 2 usage · 3 server not running · 4 conflict (409) · 5 file, DRM, path or sheet error.
