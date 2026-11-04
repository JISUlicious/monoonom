# Vendored dependencies

Pure-Python packages bundled so the skill runs without `pip install` (offline / locked-down PCs).

| Package | Version | License |
|---|---|---|
| openpyxl | 3.1.5 | MIT |
| et_xmlfile | 2.0.0 | MIT |
| xlrd | 2.0.2 | BSD |
| xlwt | 1.3.0 | BSD (with an acknowledgment clause); parts of Utils.py LGPL-2.1+ |

License texts are in `licenses/`; notices and obligations: `../../THIRD_PARTY_NOTICES.md`. `xl.py` puts this directory first on `sys.path`.
