// Server-side helper running the browser's engine on Node.js 18+.
// stdin (a) a snapshot JSON                       -> {values: [[sheetId, r, c, v]]}  (formula results)
//       (b) {snapshot, display: [[sid, r1, c1, r2, c2], ...]}
//                                                 -> {values, display: [[sid, r, c, text], ...]}
//           display = the text shown in the grid (number formats applied), as filters compare it.
import { Workbook } from '../../assets/web/js/model.js';
import { keyRow, keyCol, key } from '../../assets/web/js/address.js';
import { formatValue } from '../../assets/web/js/numfmt.js';

let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (d) => { input += d; });
process.stdin.on('end', () => {
  const req = JSON.parse(input);
  const snap = req.snapshot || req;
  const wb = Workbook.fromJSON(snap);
  const values = [];
  for (const sh of wb.sheets) {
    for (const [k, cell] of sh.cells) {
      if (!cell.f) continue;
      let v = wb.value(sh, keyRow(k), keyCol(k));
      if (v && typeof v === 'object') v = { err: v.err };
      values.push([sh.id, keyRow(k), keyCol(k), v ?? null]);
    }
  }
  const display = [];
  for (const [sid, r1, c1, r2, c2] of req.display || []) {
    const sh = wb.sheets.find((s) => s.id === sid);
    if (!sh) continue;
    for (let r = r1; r <= Math.min(r2, sh.maxRow); r++) {
      for (let c = c1; c <= Math.min(c2, sh.maxCol); c++) {
        const cell = sh.cells.get(key(r, c));
        if (!cell) continue;
        const v = wb.value(sh, r, c);
        display.push([sid, r, c, v === null || v === undefined ? '' : formatValue(v, cell.s?.nf).text]);
      }
    }
  }
  process.stdout.write(JSON.stringify({ values, display, ms: wb.lastCalcMs }));
});
