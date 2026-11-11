# Known limitations

**English** | [한국어](limitations.ko.md) · [Back to README](../../README.md)

- **xls is read-only**: only computed values are read (xlrd does not give formula text). Saving writes a `.xlsx` with the same name and leaves the `.xls` alone.
- **CSV / TSV hold one sheet of values**: they open as one sheet; a plain save writes a `.xlsx` with the same name and leaves the `.csv` alone. Saving as a `.csv` name writes only the active sheet's values (UTF-8 with BOM); formulas (their results are written), formatting, other sheets, charts and pivots are not in that file.
- Conditional formatting and sparklines are **not shown but kept when saving**. Inserting rows does not move their ranges.
- Pivot tables: kept, outlined, refreshed and created (tabular form, one column field). Edits inside a pivot and inserts, deletes or sorts across one are refused. New items in the source get their rows when Excel refreshes the pivot on opening.
- Pivot charts in a file are kept, shown and can be retitled or restyled; their data comes from the pivot table, so their data range, row/column switch and scatter type cannot be changed (as in Excel). New pivot charts can be made from pivot tables with one row field (Insert → PivotChart).
- Dynamic array formulas do not spill; only the top-left value is shown.
- Functions the app cannot compute (LAMBDA, add-in and VBA functions, …) show the results stored in the file and keep their formulas when saved.
- Chart details (per-point colors, data labels, secondary axes) are rebuilt in the default style when a chart is edited.
- Cell-level "shift cells" inserts and deletes do not adjust references in other formulas (whole-row and whole-column inserts and deletes do).
- Printing: use Excel.
