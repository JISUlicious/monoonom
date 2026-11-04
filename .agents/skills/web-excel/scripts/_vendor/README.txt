# Vendored dependencies

Pure-Python packages bundled so the skill runs without `pip install` (offline / locked-down PCs).

| Package | Version | License |
|---|---|---|
| openpyxl | 3.1.5 | MIT |
| et_xmlfile | 2.0.0 | MIT |
| xlrd | 2.0.2 | BSD-style (one part with an acknowledgment clause) |

License texts are in `licenses/`; notices and obligations: `../../THIRD_PARTY_NOTICES.md`. `xl.py` puts this directory first on `sys.path`.
