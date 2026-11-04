# Third-party notices

web-excel itself is MIT licensed (`LICENSE.txt`). It bundles the following components, unmodified,
each under its own license. Full texts are in `scripts/_vendor/licenses/` and `assets/web/vendor/`.

| Component | Version | Where | License | Full text |
|---|---|---|---|---|
| openpyxl | 3.1.5 | `scripts/_vendor/openpyxl` | MIT | `openpyxl-3.1.5-LICENCE.rst` |
| et_xmlfile | 2.0.0 | `scripts/_vendor/et_xmlfile` | MIT; includes code under the Python Software Foundation License | `et_xmlfile-2.0.0-LICENCE.rst`, `et_xmlfile-2.0.0-LICENCE.python` |
| xlrd | 2.0.2 | `scripts/_vendor/xlrd` | BSD-style (John Machin / Lingfo; David Giffin, with an acknowledgment clause) | `xlrd-2.0.2-LICENSE` |
| Chart.js | 4.4.6 | `assets/web/vendor/chart.umd.min.js` | MIT | `assets/web/vendor/chart.js.LICENSE.txt` |

## Required acknowledgment (part of xlrd)

This product includes software developed by David Giffin <david@giffin.org>.

This sentence must stay in every redistribution, and in advertising materials that mention
features or use of xlrd (xlrd-2.0.2-LICENSE, second license, clauses 3–4).

No component is under the GPL or LGPL. (xlwt, which has LGPL parts, was removed: .xls files are
read only and saved as .xlsx.)
