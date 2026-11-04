"""Python read/write round-trip tests:  python -m unittest tests/test_roundtrip.py"""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts", "_vendor"))
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "web-excel", "scripts"))

from server.app import blank_workbook, read_workbook, write_workbook  # noqa: E402
from server.formula_compat import add_prefixes, strip_prefixes  # noqa: E402
from tests.make_samples import make_xls, make_xlsx  # noqa: E402


class RoundTrip(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        cls.xlsx = os.path.join(cls.tmp, "s.xlsx")
        cls.xls = os.path.join(cls.tmp, "s.xls")
        make_xlsx(cls.xlsx)
        make_xls(cls.xls)

    def cells(self, sheet):
        return {(r, c): it for r, c, it in sheet["cells"]}

    def test_read_xlsx(self):
        m = read_workbook(self.xlsx)
        s = m["sheets"][0]
        cells = self.cells(s)
        self.assertEqual(cells[(1, 4)]["f"], "=SUM(B2:D2)")
        self.assertEqual(cells[(11, 1)]["f"], '=XLOOKUP("3월",A2:A7,C2:C7)')  # _xlfn stripped
        self.assertEqual(s["freeze"], "B2")
        self.assertEqual(s["autoFilter"]["ref"], "A1:F7")
        self.assertEqual(s["charts"][0]["type"], "bar")
        self.assertEqual(s["charts"][0]["series"][-1]["type"], "line")  # combo chart
        self.assertEqual(m["styles"][cells[(0, 0)]["s"]]["bg"], "#4472C4")
        self.assertEqual(m["sheets"][1]["merges"], ["A6:B6"])

    def test_xlsx_patch_roundtrip(self):
        m = read_workbook(self.xlsx)
        m["sheets"][0]["name"] = "판매"
        m["sheets"][0]["cells"].append([20, 0, {"v": "추가"}])
        out = os.path.join(self.tmp, "o.xlsx")
        write_workbook(out, "xlsx", m, self.xlsx)
        m2 = read_workbook(out)
        self.assertEqual(m2["sheets"][0]["name"], "판매")
        self.assertEqual(self.cells(m2["sheets"][0])[(20, 0)]["v"], "추가")
        self.assertEqual(len(m2["sheets"][0]["charts"]), 1)
        self.assertEqual(len(m2["sheets"][1]["charts"]), 1)

    def test_new_chart_written(self):
        m = blank_workbook()
        s = m["sheets"][0]
        s["cells"] = [[0, 0, {"v": "a"}], [0, 1, {"v": 1}], [1, 0, {"v": "b"}], [1, 1, {"v": 2}]]
        s["charts"] = [{"type": "pie", "title": "T", "legend": "r", "dirty": True,
                        "series": [{"cat": "Sheet1!$A$1:$A$2", "val": "Sheet1!$B$1:$B$2"}],
                        "anchor": {"c": 3, "r": 1, "w": 300, "h": 200}}]
        out = os.path.join(self.tmp, "n.xlsx")
        write_workbook(out, "xlsx", m, None)
        c = read_workbook(out)["sheets"][0]["charts"][0]
        self.assertEqual((c["type"], c["title"], c["anchor"]["w"]), ("pie", "T", 300))

    def test_xls_read_write(self):
        m = read_workbook(self.xls)
        self.assertEqual(self.cells(m["sheets"][0])[(1, 0)]["v"], "사과")
        # xlsx model -> xls: formulas keep computed (cached) values
        x = read_workbook(self.xlsx)
        for r, c, it in x["sheets"][0]["cells"]:
            if (r, c) == (7, 1):
                it["cv"] = 951  # value computed by the browser engine
        out = os.path.join(self.tmp, "o.xls")
        write_workbook(out, "xls", x, None)
        cells = self.cells(read_workbook(out)["sheets"][0])
        self.assertEqual(cells[(7, 1)]["v"], 951)
        self.assertEqual(cells[(0, 0)]["v"], "월")

    def test_prefixes(self):
        self.assertEqual(add_prefixes('=XLOOKUP(A1,B:B,C:C)+SUM(1)'), '=_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(1)')
        self.assertEqual(add_prefixes('=FILTER(A1:A3,B1:B3)'), '=_xlfn._xlws.FILTER(A1:A3,B1:B3)')
        self.assertEqual(add_prefixes('="XLOOKUP("'), '="XLOOKUP("')  # not inside strings
        self.assertEqual(strip_prefixes('=_xlfn.IFS(A1,1)'), '=IFS(A1,1)')


if __name__ == "__main__":
    unittest.main()
