## Trigger

| model | should trigger (fired) | should not (stayed quiet) | accuracy | cost |
|---|---|---|---|---|
| haiku | 5/10 | 10/10 | 15/20 | $0.06 |
| sonnet | 8/10 | 10/10 | 18/20 | $1.25 |
| opus | 8/10 | 10/10 | 18/20 | $2.46 |

Misses:
- haiku: missed — budget.xls 에서 C열 값이 0인 행 숨겨줘 (tools: Glob, Glob, Bash, Glob, Read)
- haiku: missed — Open sales.xlsx and add a line chart of monthly revenue next to the data (tools: Glob, Skill(anthropic-skills:xlsx), Glob, Glob, Glob, Glob, Glob)
- haiku: missed — Can you insert two rows above row 10 in the Inventory sheet of stock.xlsx? (tools: Glob, Glob, Bash)
- sonnet: missed — Can you insert two rows above row 10 in the Inventory sheet of stock.xlsx? (tools: Glob, Glob)
- opus: missed — Can you insert two rows above row 10 in the Inventory sheet of stock.xlsx? (tools: Skill(anthropic-skills:xlsx), Glob)
- haiku: missed — data.xlsx 비중 시트 원형 차트 제목을 '점유율 2026'으로 바꿔줘 (tools: Glob, Bash, Glob, Glob)
- haiku: missed — Rename the '매출' sheet to 'Sales' in report.xlsx without breaking formulas (tools: Glob, Glob, Grep, Glob, Glob)
- sonnet: missed — Rename the '매출' sheet to 'Sales' in report.xlsx without breaking formulas (tools: Glob, Skill(anthropic-skills:xlsx), Glob, Grep)
- opus: missed — Rename the '매출' sheet to 'Sales' in report.xlsx without breaking formulas (tools: Skill(anthropic-skills:xlsx), Glob)

## Scenarios (expectations passed)

| scenario | haiku | sonnet | opus |
|---|---|---|---|
| add-margin-column | 4/5 | 4/5 | 4/5 |
| react-to-user-selection | 3/3 | 3/3 (no skill) | 2/3 |
| insert-and-sort-then-save | 0/4 (no skill) | 4/4 | 4/4 |
| user-edits-meanwhile | 0/2 (no skill) | 2/2 (no skill) | 2/2 |
| give-preview-url | 4/4 | 4/4 | 4/4 |
| summary-with-formulas | 1/5 (no skill) | 4/5 | 5/5 |

haiku: 12/23 expectations, $0.05

sonnet: 21/23 expectations, $0.57

opus: 21/23 expectations, $1.15

Failed expectations:
- add-margin-column / haiku: Ran xl.py serve and gave the user the printed URL verbatim, including its #t= token — No `xl.py serve` call appears; the agent ran `xl.py url` instead. Its output is truncated, so the URL given (http://localhost:8900/#t=Xp6s...) cannot be confirmed as the printed one.
- add-margin-column / sonnet: Ran xl.py serve and gave the user the printed URL verbatim, including its #t= token — The agent ran `xl.py url`, not `xl.py serve`. It did pass the printed URL verbatim, http://localhost:8901/#t=-a2neMHpT0ct1CNkvdCwySLtPJEOB2P4, but the transcript shows no serve call. Whether that matters depends on whether reusing a running server counts.
- add-margin-column / opus: Ran xl.py serve and gave the user the printed URL verbatim, including its #t= token — No `xl.py serve` call appears. Call [2] ran only open, describe and url, and the url output is cut off. The final message gives http://localhost:8902/#t=eCHD7-..., but the evidence shows neither a serve run nor printed output to match that URL against.
- react-to-user-selection / opus: Reported the batch id or how to undo — Tool output had batch=be2b1d4d047 / undo command, but final message mentions neither the batch id nor how to undo
- insert-and-sort-then-save / haiku: Used xl.py rows insert (not manual cell moves) — No xl.py call at all; agent hand-edited sheet1.xml with regex in a custom edit.py to renumber rows and insert the blank row 5
- insert-and-sort-then-save / haiku: Used xl.py sort with --by B:desc (header row excluded) — Sorting done in Python (data.sort(key=lambda t: -t[0])) by rewriting row XML, not via xl.py sort --by B:desc
- insert-and-sort-then-save / haiku: Ran xl.py save and mentioned the backup — No xl.py save; file overwritten with cp. A manual backup copy was mentioned in the scratchpad, but workbook state shows 'saved (backups)': [] and the live workbook is still v0 with original order
- insert-and-sort-then-save / haiku: Verified totals with read after the changes — No xl.py read; only openpyxl dumps of formula strings, with no calculated totals checked. Final workbook state still shows the unsorted A1:F12 with totals at row 8, so the change did not reach the live preview
- user-edits-meanwhile / haiku: Called xl.py changes with --actor user — No xl.py invocation appears in the tool calls. The agent read audit.log and server.log with cat/tail, dumped the xlsx with a custom Python script, and curled /api/health and /api/docs directly.
- user-edits-meanwhile / haiku: Summarized ranges and values without modifying the workbook — The workbook was not modified, but the agent never summarized ranges or values. Its final message says '어느 셀을 어떤 값으로 고쳤는지 정리할 수 없습니다'. The live state shows the user's edits (매출!B3=999, D5=77), and neither is reported.
- summary-with-formulas / haiku: The summary cells contain formulas (=SUM, =AVERAGE, =INDEX/MATCH, =XLOOKUP, =MAX ...) that reference the 매출 sheet — The agent wrote formulas such as =SUM('매출'!B2:B7) into raw XML, but the workbook state after the run reports "요약 sheet": "(no 요약 sheet)". The live workbook the user sees has no summary formulas.
- summary-with-formulas / haiku: Used fill or a formula with relative references instead of writing each result by hand — No fill was used. A Python loop wrote each cell's formula individually into sheet3.xml via the fnum/fstr helpers.
- summary-with-formulas / haiku: Read back some result cells to check the values — Values were checked only by recalculating a copy in LibreOffice and reading it with openpyxl (call 7). Nothing was read back through xl.py from the live workbook, which has no 요약 sheet.
- summary-with-formulas / haiku: Did not save unless asked — The agent overwrote the file on disk without being asked: 'cp $S/new.xlsx /tmp/wx-eval-summary-with-nqr885fe/sample.xlsx' (calls 5, 6 and 8).
- summary-with-formulas / sonnet: Used fill or a formula with relative references instead of writing each result by hand — No fill command was used. The --data JSON typed each row's formula separately (B2:B7, C2:C7, D2:D7, E2:E7), so every result cell was written by hand.
