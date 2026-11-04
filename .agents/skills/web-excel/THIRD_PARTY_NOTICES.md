# Third-party notices

web-excel itself is MIT licensed (`LICENSE.txt`). It bundles the following components, unmodified,
each under its own license. Full texts are in `scripts/_vendor/licenses/` and `assets/web/vendor/`.

| Component | Version | Where | License | Full text |
|---|---|---|---|---|
| openpyxl | 3.1.5 | `scripts/_vendor/openpyxl` | MIT | `openpyxl-3.1.5-LICENCE.rst` |
| et_xmlfile | 2.0.0 | `scripts/_vendor/et_xmlfile` | MIT; includes code under the Python Software Foundation License | `et_xmlfile-2.0.0-LICENCE.rst`, `et_xmlfile-2.0.0-LICENCE.python` |
| xlrd | 2.0.2 | `scripts/_vendor/xlrd` | BSD-style (two notices) | `xlrd-2.0.2-LICENSE` |
| xlwt | 1.3.0 | `scripts/_vendor/xlwt` | BSD-style (John Machin; pyExcelerator); parts of `xlwt/Utils.py` from pyXLWriter under LGPL-2.1-or-later | `xlwt-1.3.0-LICENSE.txt`, `LGPL-2.1.txt` |
| Chart.js | 4.4.6 | `assets/web/vendor/chart.umd.min.js` | MIT | `assets/web/vendor/chart.js.LICENSE.txt` |

## Required acknowledgment (pyExcelerator, part of xlwt)

This product includes software developed by Roman V. Kiseliov <roman@kiseliov.ru>.

## LGPL component (parts of xlwt.Utils)

xlwt is shipped as unmodified Python source, so you can read, change or replace it
(`scripts/_vendor/xlwt`, or install another xlwt and remove the bundled copy). web-excel only
uses it as a library: `scripts/server/xls_io.py` overrides one method (`FormulaCell.get_biff_data`,
to store cached formula results) at run time without changing xlwt's files; that override is also
available under the LGPL-2.1-or-later. Upstream source: https://github.com/python-excel/xlwt
(release 1.3.0, https://pypi.org/project/xlwt/1.3.0/).
