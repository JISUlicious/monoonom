// Formula engine unit tests: node tests/engine.test.mjs
import { blankWorkbook, Sheet } from '../.agents/skills/monoonom/assets/web/js/model.js';
import { translateFormula, adjustFormula, cycleAbsolute } from '../.agents/skills/monoonom/assets/web/js/formula.js';
import { formatValue } from '../.agents/skills/monoonom/assets/web/js/numfmt.js';

const wb = blankWorkbook();
const s = wb.sheet;
const put = (ref, val, sh = s) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, r = +m[2] - 1;
  wb.setCell(sh, r, c, typeof val === 'string' && val.startsWith('=') ? { f: val } : { v: val });
};
const data = [['이름', '수량', '단가'], ['사과', 3, 1200], ['배', 2, 2500], ['포도', 5, 3000], ['사과', 1, 1100]];
data.forEach((row, r) => row.forEach((v, c) => wb.setCell(s, r, c, { v })));
const s2 = wb.addSheet('데이터 2'); wb.active = 0;
wb.setCell(s2, 0, 0, { v: 10 });
const cases = [
  ['=SUM(B2:B5)', 11], ['=B2*C2+B3*C3', 8600], ['=SUMPRODUCT(B2:B5,C2:C5)', 24700],
  ['=SUMIF(A2:A5,"사과",B2:B5)', 4], ['=COUNTIF(B2:B5,">2")', 2], ['=SUMIFS(B2:B5,A2:A5,"사과",C2:C5,">1150")', 3],
  ['=AVERAGE(B2:B5)', 2.75], ['=MAX(C:C)', 3000], ['=VLOOKUP("배",A2:C5,3,FALSE)', 2500],
  ['=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))', '포도'], ['=XLOOKUP("포도",A2:A5,C2:C5)', 3000],
  ['=IFERROR(1/0,"err")', 'err'], ['=1/0', '#DIV/0!'], ['=-2^2', 4], ['=2^3^2', 64], ['=10%', 0.1],
  ['="a"&1&TRUE', 'a1TRUE'], ['=IF(B2>2,"큼","작음")', '큼'], ['=LEFT("가나다라",2)', '가나'],
  ["='데이터 2'!A1*2", 20], ['=TEXT(1234.5,"#,##0.0")', '1,234.5'], ['=DATE(2026,2,30)', 46083],
  ['=TEXT(DATE(2026,1,1),"yyyy-mm-dd")', '2026-01-01'], ['=EOMONTH(DATE(2026,1,15),1)', 46081],
  ['=DATEDIF(DATE(2020,1,1),DATE(2026,3,1),"Y")', 6], ['=ROUND(2.345,2)', 2.35], ['=ROUND(-2.5,0)', -3],
  ['=MOD(-3,2)', 1], ['=COUNTA(A1:C5)', 15], ['=COUNTBLANK(A1:D5)', 5], ['=ROWS(A1:B9)', 9],
  ['=SUM(A1:INDEX(B1:B5,3))', 5], ['=CONCAT(A2:A3)', '사과배'], ['=TEXTJOIN(",",TRUE,A2:A5)', '사과,배,포도,사과'],
  ['=SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'], ['=PMT(0.05/12,12,1000000)', -85607.4817885],
  ['=NPV(0.1,-100,50,60)', -100 / 1.1 + 50 / 1.21 + 60 / 1.331], ['=SUM(OFFSET(B1,1,0,4,1))', 11],
  ['=INDIRECT("B"&3)', 2], ['=AND(TRUE,1)', true], ['=SUM({1,2;3,4})', 10], ['=SUMPRODUCT((A2:A5="사과")*B2:B5)', 4],
  ['=MEDIAN(1,2,3,4)', 2.5], ['=STDEV(2,4,4,4,5,5,7,9)', 2.138089935], ['=LARGE(C2:C5,2)', 2500],
  ['=RANK(2500,C2:C5)', 2], ['=WEEKDAY(DATE(2026,11,4))', 4], ['=0.1+0.2=0.3', true],
  ['=UNKNOWNFN(1)', '#NAME?'], ['=A1:A3', '이름'], ['=SUBTOTAL(9,B2:B5)', 11], ['=SWITCH(2,1,"a",2,"b")', 'b'],
  ['=IFS(B2>5,"x",B2>1,"y")', 'y'], ['=ISBLANK(Z99)', true], ['=LEN(A2)', 2], ['=VALUE("1,234")', 1234],
  ['=MATCH(2600,C2:C4,1)', 2], ['=HLOOKUP("단가",A1:C5,2,FALSE)', 1200], ['=LET(x,2,y,3,x*y)', 6],
];
let fail = 0;
for (const [f, exp] of cases) {
  const v = wb.engine.evalFormula(f, s, 9, 9);
  const got = v && v.err ? v.err : v;
  const ok = typeof exp === 'number' ? Math.abs(got - exp) < 1e-6 : got === exp;
  if (!ok) { fail++; console.log('FAIL', f, '=>', got, 'expected', exp); }
}
// chain recalculation (forward and reverse 20k-deep chains)
const N = 20000;
wb.begin('t'); for (let r = 0; r < N; r++) wb.setCell(s, r, 5, r === 0 ? { v: 1 } : { f: `=F${r}+1` }); wb.commit();
if (wb.value(s, N - 1, 5) !== N) { fail++; console.log('FAIL chain', wb.value(s, N - 1, 5)); }
wb.begin('t'); for (let r = 0; r < N; r++) wb.setCell(s, r, 7, r === N - 1 ? { v: 1 } : { f: `=H${r + 2}+1` }); wb.commit();
if (wb.value(s, 0, 7) !== N) { fail++; console.log('FAIL reverse chain', wb.value(s, 0, 7)); }
// circular reference -> 0, recorded
wb.begin('t'); wb.setCell(s, 0, 10, { f: '=L1+1' }); wb.setCell(s, 0, 11, { f: '=K1+1' }); wb.commit();
if (!wb.engine.circular.size) { fail++; console.log('FAIL circular'); }
wb.begin('t'); wb.setCell(s, 0, 10, null); wb.setCell(s, 0, 11, null); wb.commit();
const tr = [
  [translateFormula('=A1+$B$2+B$3+$C4', 2, 1), '=B3+$B$2+C$3+$C6'],
  [adjustFormula('=SUM(A1:A10)+B5', 'Sheet1', 'Sheet1', 'row', 2, 3), '=SUM(A1:A13)+B8'],
  [adjustFormula('=SUM(A1:A10)+B5', 'Sheet1', 'Sheet1', 'row', 4, -1), '=SUM(A1:A9)+#REF!'],
  [adjustFormula("='데이터 2'!A1+A1", 'Sheet1', '데이터 2', 'col', 0, 1), "='데이터 2'!B1+A1"],
  [cycleAbsolute('=A1+B2', 2).text, '=$A$1+B2'],
  [translateFormula('=A1', -1, 0), '=#REF!'],
];
for (const [g, e] of tr) if (g !== e) { fail++; console.log('FAIL rewrite', g, 'expected', e); }
// undo / redo
const before = wb.value(s, 1, 1);
wb.setCell(s, 1, 1, { v: 99 });
wb.undo();
if (wb.value(s, 1, 1) !== before) { fail++; console.log('FAIL undo'); }
wb.redo();
if (wb.value(s, 1, 1) !== 99) { fail++; console.log('FAIL redo'); }
// insert row shifts formulas
wb.setCell(s, 20, 0, { f: '=SUM(B2:B5)' });
wb.shift(s, 'row', 0, 1);
if (s.get(21, 0)?.f !== '=SUM(B3:B6)') { fail++; console.log('FAIL shift', s.get(21, 0)); }
console.log(`${cases.length + tr.length + 5} checks, ${fail} failed`);
process.exit(fail ? 1 : 0);
