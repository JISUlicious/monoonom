# Third-party notices

web-excel itself is MIT licensed (`LICENSE.txt`). The only third-party code shipped in this folder is:

| Component | Version | Where | License | Full text |
|---|---|---|---|---|
| Chart.js | 4.4.6 | `assets/web/vendor/chart.umd.min.js` | MIT | `assets/web/vendor/chart.js.LICENSE.txt` (also the header of the file) |

It is bundled because the preview page must work without internet access (no CDN).

Python packages are **not** bundled; they are installed by the user from PyPI with
`python -m pip install -r scripts/requirements.txt` and keep their own licenses there:
openpyxl and its dependency et_xmlfile (MIT), and optionally xlrd (BSD-style, for .xls).
