# Batch edit format (`xl.py apply`)

A JSON array of edits, applied in order as one batch (one undo):

```json
[
  {"op": "set", "range": "Sheet1!A1", "values": [["Month", "Sales"], ["Jan", 120], ["Feb", 135]]},
  {"op": "set", "range": "C2:C3", "values": [["=B2*1.1"]]},
  {"op": "format", "range": "A1:C1", "set": {"b": true, "bg": "#DDEBF7"}, "border": {"kind": "bottom", "style": "medium"}},
  {"op": "chart", "data": "A1:B3", "type": "bar", "title": "Sales", "at": "E2"}
]
```

| op | fields |
|---|---|
| `set` | `range`, `values` (2-D array, 1-D row, or one value to fill the range), `parse` (default true). A cell value may also be `{"f": "=..."}`, `{"text": "..."}`, `{"date": "2026-03-15"}` or `{"v": 1, "nf": "0.0%"}`. Filling one formula over a range shifts relative references. |
| `clear` | `range`, `what`: contents \| formats \| all |
| `format` | `range`, `set` (style keys), `unset` (list), `border` (`"outline"` or `{"kind","style","color"}`) |
| `merge` / `unmerge` | `range`, `center` (default true) |
| `colWidth` | `range` (e.g. `"B:D"`), `width` in characters |
| `rowHeight` | `range` (e.g. `"1:3"`), `height` in points |
| `hide` / `unhide` | `range` of columns (`"B:C"`) or rows (`"3:5"`) |
| `freeze` | `sheet`, `cell` (`null` to unfreeze) |
| `filter` | `range` (`null` to remove), `sheet` |
| `chart` | `data`, `type`, `title`, `at`, `sheet`, `legend`, `width`, `height`, `xTitle`, `yTitle`, `byRows` |
| `deleteChart` | `sheet`, `id` or `index` |
| `addSheet` | `name`, `at` |
| `renameSheet` | `sheet`, `name` |
| `deleteSheet` | `sheet` |
| `name` | `name`, `ref` (omit `ref` to delete a defined name) |

Style keys: `b`, `i`, `u` (true or "double"), `st`, `fs`, `ff`, `fc`, `bg`, `ha`, `va`, `wrap`, `ind`,
`rot`, `nf`, and borders `bt`/`bb`/`bl`/`br` as `{"s": "thin", "c": "#000000"}`. A key set to `null`
or `false` removes it.

Ranges may include a sheet (`'My Sheet'!A1:B2`); without one, the workbook's active sheet is used
unless the edit has `"sheet"`.
