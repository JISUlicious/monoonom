"""Python read/write round-trip tests:  python -m unittest tests/test_roundtrip.py"""
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, ".agents", "skills", "monoonom", "scripts"))

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
        try:
            make_xls(cls.xls)
        except FileNotFoundError as e:  # only the .xls test depends on the fixture
            cls.xls, cls.xls_missing = None, str(e)

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

    def test_xls_read_only(self):
        if not self.xls:
            self.skipTest(self.xls_missing)
        m = read_workbook(self.xls)
        cells = self.cells(m["sheets"][0])
        self.assertEqual(cells[(1, 0)]["v"], "사과")
        self.assertEqual(cells[(1, 2)]["v"], 1200)
        self.assertTrue(any(".xlsx" in w for w in m["warnings"]))
        # .xls is never written; targets are converted to .xlsx
        from server.app import ApiError, xlsx_target
        with self.assertRaises(ApiError):
            write_workbook(os.path.join(self.tmp, "o.xls"), "xls", m, None)
        self.assertEqual(xlsx_target("/w/report.xls", "xls"), ("/w/report.xlsx", "xlsx", True))
        self.assertEqual(xlsx_target("/w/report.xlsx", "xlsx"), ("/w/report.xlsx", "xlsx", False))
        # a stale client asking for 'xls' on an xlsx/xlsm path: keep the path, no double extension
        self.assertEqual(xlsx_target("/w/report.xlsx", "xls"), ("/w/report.xlsx", "xlsx", False))
        self.assertEqual(xlsx_target("/w/report.xlsm", "xls"), ("/w/report.xlsm", "xlsm", False))
        # an explicit xlsm for an .xls path is kept
        self.assertEqual(xlsx_target("/w/out.xls", "xlsm"), ("/w/out.xlsm", "xlsm", True))
        self.assertEqual(xlsx_target("/w/out.XLS", None), ("/w/out.xlsx", "xlsx", True))
        out = os.path.join(self.tmp, "from-xls.xlsx")
        write_workbook(out, "xlsx", m, self.xls)  # base .xls is ignored, xlsx written from the model
        self.assertEqual(self.cells(read_workbook(out)["sheets"][0])[(0, 3)]["v"], "금액")

    def test_prefixes(self):
        self.assertEqual(add_prefixes('=XLOOKUP(A1,B:B,C:C)+SUM(1)'), '=_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(1)')
        self.assertEqual(add_prefixes('=FILTER(A1:A3,B1:B3)'), '=_xlfn._xlws.FILTER(A1:A3,B1:B3)')
        self.assertEqual(add_prefixes('="XLOOKUP("'), '="XLOOKUP("')  # not inside strings
        self.assertEqual(strip_prefixes('=_xlfn.IFS(A1,1)'), '=IFS(A1,1)')


if __name__ == "__main__":
    unittest.main()
