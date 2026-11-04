// Server-side recalculation with the same engine the browser uses (needs Node.js 18+).
// stdin: document snapshot JSON (GET /api/docs/{id}/snapshot)  stdout: {values: [[sheetId, r, c, v]]}
import { Workbook } from '../../assets/web/js/model.js';
import { keyRow, keyCol } from '../../assets/web/js/address.js';

let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (d) => { input += d; });
process.stdin.on('end', () => {
  const snap = JSON.parse(input);
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
  process.stdout.write(JSON.stringify({ values, ms: wb.lastCalcMs }));
});
