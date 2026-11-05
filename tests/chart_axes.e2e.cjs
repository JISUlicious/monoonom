// Value axes as Excel scales them, axis labels in the cells' format, and series whose values are
// stored in the chart itself (numLit / strLit). Starts its own server (port 8844).
//   node tests/chart_axes.e2e.cjs
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'monoonom', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-axes-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });

py(`
import datetime as dt
from openpyxl import Workbook
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.chart.data_source import NumDataSource, NumData, NumVal, AxDataSource, StrData, StrVal
wb = Workbook(); sh = wb.active; sh.title = 'S'
rows = [['키', '값', '혼합', '날짜', '측정'], ['a', 160, -3], ['b', 170, 5], ['c', 180, 2], ['d', 175, 4]]
for r in rows: sh.append(r)
for i, d in enumerate([dt.date(2026, 1, 1), dt.date(2026, 2, 1), dt.date(2026, 3, 1), dt.date(2026, 4, 1)]):
    sh.cell(i + 2, 4, d).number_format = 'yyyy-mm-dd'; sh.cell(i + 2, 5, [3, 5, 4, 6][i])
b = BarChart(); b.title = 'narrow'; b.add_data(Reference(sh, min_col=2, min_row=1, max_row=5), titles_from_data=True)
b.set_categories(Reference(sh, min_col=1, min_row=2, max_row=5)); sh.add_chart(b, 'G2')
m = BarChart(); m.title = 'mixed'; m.add_data(Reference(sh, min_col=3, min_row=1, max_row=5), titles_from_data=True); sh.add_chart(m, 'G20')
s = ScatterChart(); s.title = 'dates'
s.series.append(Series(Reference(sh, min_col=5, min_row=2, max_row=5), Reference(sh, min_col=4, min_row=2, max_row=5), title='측정'))
sh.add_chart(s, 'P2')
l = BarChart(); l.title = 'literal'
ser = Series(Reference(sh, min_col=2, min_row=2, max_row=5), title='목록')
ser.val = NumDataSource(numLit=NumData(ptCount=3, pt=[NumVal(idx=0, v=7), NumVal(idx=1, v=9), NumVal(idx=2, v=8)]))
ser.cat = AxDataSource(strLit=StrData(ptCount=3, pt=[StrVal(idx=0, v='가'), StrVal(idx=1, v='나'), StrVal(idx=2, v='다')]))
l.series.append(ser); sh.add_chart(l, 'P20')
wb.save(${JSON.stringify(path.join(ws, 'axes.xlsx'))})
`);

(async () => {
  xl('serve', '--port', '8844'); xl('open', 'axes.xlsx');
  const ls = xl('chart', 'ls');
  check('chart ls: a series stored in the chart shows its values', /'literal'[^\n]*\n\s+계열 1: 이름 '목록'\s+X \{가, 나, 다\} \(차트에 저장된 값 3개\)\s+Y \{7, 9, 8\} \(차트에 저장된 값 3개\)/.test(ls), ls);
  const browser = await chromium.launch();
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.wb.sheet.charts.length === 4); await sleep(600);
    const ax = await p.evaluate(() => {
      const a = window.monoonom, sh = a.doc.wb.sheet, out = {};
      for (const m of sh.charts) {
        const ch = a.charts.instances.get(m.id).chart, y = ch.scales.y, x = ch.scales.x;
        out[m.title] = { yMin: y.min, yMax: y.max, xMin: x.min, xMax: x.max, xTicks: x.ticks.map((t) => t.label), yTicks: y.ticks.map((t) => t.label), data: ch.data.datasets[0].data, labels: ch.data.labels };
      }
      return out;
    });
    const n = ax.narrow, mx = ax.mixed, d = ax.dates, lit = ax.literal;
    check('160–180 (spread under 1/6): the axis does not start at 0, but below 160 (Excel 5/6 rule)', n.yMin > 0 && n.yMin < 160 && n.yMax > 180, JSON.stringify([n.yMin, n.yMax]));
    // the 5/6 boundary with Peltier's numbers: min exactly 5/6 of max -> not 0; just under -> 0
    const edge = await p.evaluate(async () => { const { excelAutoRange: f } = await import('/js/charts.js');
      return [f([2.5, 3]).beginAtZero, f([2.49999999999999, 3]).beginAtZero, f([-3, -2.5]).beginAtZero, f([-3, -2.49999999999999]).beginAtZero]; });
    check('5/6 rule boundary: exactly 5/6 does not start at 0, just under does (both signs)', edge.join() === 'false,true,false,true', edge.join());
    check('mixed signs: 0 inside the axis', mx.yMin < -3 && mx.yMax > 5, JSON.stringify([mx.yMin, mx.yMax]));
    check('scatter with dates as X: X labels in the cells\' date format', d.xTicks.length > 1 && d.xTicks.every((t) => /^\d{4}-\d{2}-\d{2}$/.test(t)), JSON.stringify(d.xTicks));
    check('…X axis does not start at 0 (1900) but near the first date (5/6 rule on X)', d.xMin > 46000, String(d.xMin));
    check('values stored in the chart are drawn (7, 9, 8 at 가, 나, 다)', JSON.stringify(lit.data) === '[7,9,8]' && lit.labels.join() === '가,나,다', JSON.stringify(lit));
    await p.evaluate(() => { const a = window.monoonom; a.charts.select(a.doc.wb.sheet.charts.find((m) => m.title === 'literal').id); }); await sleep(200);
    const panel = await p.evaluate(() => document.querySelector('.chart-ranges')?.textContent || '');
    check('range panel: the stored values are listed (no cells to outline)', panel.includes('Y(값) {7, 9, 8}') && panel.includes('차트에 저장된 값'), panel);
    if (process.env.SHOT) await p.screenshot({ path: process.env.SHOT });
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  // saved file: the edited literal chart keeps its values; value / scatter axes link to the cells' format
  const id = ls.split('\n').find((l) => l.includes("'literal'")).split(/\s+/)[1];
  xl('chart', 'set', id, '--title', 'literal2');
  const did = ls.split('\n').find((l) => l.includes("'dates'")).split(/\s+/)[1];
  xl('chart', 'set', did, '--title', 'dates2');
  xl('save', '--as', 'out.xlsx');
  const r = JSON.parse(py(`
import json, re, zipfile
z = zipfile.ZipFile(${JSON.stringify(path.join(ws, 'out.xlsx'))}); out = {}
for n in z.namelist():
    if n.startswith('xl/charts/chart'):
        s = z.read(n).decode(); t = re.findall(r'<a:t>([^<]*)', s); out[t[0] if t else n] = s
lit, sc = out['literal2'], out['dates2']
print(json.dumps({'numLit': re.findall(r'<numLit>.*?</numLit>', lit, re.S), 'strLit': re.findall(r'<strLit>.*?</strLit>', lit, re.S),
  'litNumRef': '<numRef>' in re.sub(r'<tx>.*?</tx>', '', lit, flags=re.S),
  'scatterFmt': re.findall(r'<valAx>.*?<numFmt formatCode="([^"]*)" sourceLinked="([^"]*)"', sc, re.S)}))`));
  check('saved: the stored values and categories are written back as numLit / strLit (no range made up)',
    r.numLit.length === 1 && /v>7</.test(r.numLit[0]) && /v>8</.test(r.numLit[0]) && r.strLit.length === 1 && r.strLit[0].includes('다') && !r.litNumRef, JSON.stringify(r));
  check('saved: both scatter axes have numFmt sourceLinked=1 (dates shown as dates in Excel)', r.scatterFmt.length === 2 && r.scatterFmt.every(([, l]) => l === '1'), JSON.stringify(r.scatterFmt));
  if (process.env.KEEP) fs.copyFileSync(path.join(ws, 'out.xlsx'), process.env.KEEP);
  try { xl('stop', '--force'); } catch {}
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch {} process.exit(1); });
