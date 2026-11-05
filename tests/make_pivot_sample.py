"""Make samples/pivot.xlsx: a real pivot table (rows 지역, columns 제품, values sum of 매출) written by
LibreOffice, for the pivot tests. Needs LibreOffice with Python UNO (`python3 -c "import uno"`).

    python3 tests/make_pivot_sample.py

Sheet 데이터: A1:D13 (지역, 제품, 분기, 매출). Sheet 피벗: the pivot table at A3, plus a note in A1 and a
cell under the pivot (A12) so row inserts above / below it can be tested.
"""
import os
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "pivot.xlsx")

ROWS = [["지역", "제품", "분기", "매출"],
        ["서울", "A", "Q1", 120], ["서울", "B", "Q1", 80], ["부산", "A", "Q1", 60], ["부산", "B", "Q1", 90],
        ["대구", "A", "Q1", 40], ["대구", "B", "Q1", 30], ["서울", "A", "Q2", 130], ["서울", "B", "Q2", 70],
        ["부산", "A", "Q2", 65], ["부산", "B", "Q2", 95], ["대구", "A", "Q2", 45], ["대구", "B", "Q2", 35]]


def main():
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.table import CellAddress, CellRangeAddress

    port = 2099
    prof = tempfile.mkdtemp(prefix="lo-prof-")
    lo = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{prof}",
                           f"--accept=socket,host=127.0.0.1,port={port};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = None
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (hidden,))
        sheets = doc.Sheets
        data = sheets.getByIndex(0); data.Name = "데이터"
        for r, row in enumerate(ROWS):
            for c, v in enumerate(row):
                cell = data.getCellByPosition(c, r)
                if isinstance(v, str):
                    cell.String = v
                else:
                    cell.Value = v
        sheets.insertNewByName("피벗", 1)
        piv = sheets.getByName("피벗")
        piv.getCellByPosition(0, 0).String = "지역별 제품 매출 (피벗)"
        piv.getCellByPosition(0, 11).String = "피벗 아래 메모"

        src = CellRangeAddress(); src.Sheet = 0; src.StartColumn = 0; src.StartRow = 0; src.EndColumn = 3; src.EndRow = len(ROWS) - 1
        tables = piv.getDataPilotTables()
        desc = tables.createDataPilotDescriptor()
        desc.setSourceRange(src)
        from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, DATA
        from com.sun.star.sheet.GeneralFunction import SUM
        fields = desc.getDataPilotFields()
        fields.getByIndex(0).Orientation = ROW       # 지역
        fields.getByIndex(1).Orientation = COLUMN    # 제품
        f = fields.getByIndex(3); f.Orientation = DATA; f.Function = SUM  # 매출
        dest = CellAddress(); dest.Sheet = 1; dest.Column = 0; dest.Row = 2
        tables.insertNewByName("피벗1", dest, desc)

        flt = PropertyValue(); flt.Name = "FilterName"; flt.Value = "Calc MS Excel 2007 XML"
        ow = PropertyValue(); ow.Name = "Overwrite"; ow.Value = True
        doc.storeToURL(uno.systemPathToFileUrl(OUT), (flt, ow))
        doc.close(True)
        print("wrote", OUT)
    finally:
        lo.terminate()


if __name__ == "__main__":
    sys.exit(main())
