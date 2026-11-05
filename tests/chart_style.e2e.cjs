// The preview draws charts with the colors and style values Excel gets: the workbook's theme
// accents, the chart's own series/slice colors, and server/chart_style.py's numbers.
// Starts its own server (port 8838).   node tests/chart_style.e2e.cjs
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'web-excel', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-cstyle-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

// an Office 2007-theme workbook (as openpyxl writes) with explicitly colored charts
execFileSync('python3', ['-c', `
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
wb = Workbook(); sh = wb.active; sh.title = 'S'
for r in [['m', 'A', 'B'], ['1', 3, 4], ['2', 5, 2], ['3', 4, 6]]: sh.append(r)
b = BarChart(); b.title = 'colored'; b.add_data(Reference(sh, min_col=2, max_col=3, min_row=1, max_row=4), titles_from_data=True)
b.series[0].graphicalProperties = GraphicalProperties(solidFill='C00000'); sh.add_chart(b, 'E2')
p = PieChart(); p.title = 'slices'; p.add_data(Reference(sh, min_col=2, min_row=1, max_row=4), titles_from_data=True)
p.series[0].dPt = [DataPoint(idx=0, spPr=GraphicalProperties(solidFill='112233'))]; sh.add_chart(p, 'E20')
wb.save(${JSON.stringify(path.join(ws, 'old.xlsx'))})
`]);

(async () => {
  xl('serve', '--port', '8838'); xl('open', 'old.xlsx');
  xl('new', '--name', 'fresh');
  xl('write', 'A1', '--data', '[["월","A","B","C"],["1월",3,4,5],["2월",5,2,4],["3월",4,6,3],["4월",6,5,2]]', '--doc', 'fresh');
  xl('chart', 'add', 'A1:D5', '--title', '막대', '--doc', 'fresh');
  xl('chart', 'add', 'A1:B5', '--type', 'line', '--at', 'F20', '--doc', 'fresh');
  const browser = await chromium.launch();
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.webExcel?.docs?.length === 2);
    const cfgs = (name) => p.evaluate((n) => {
      const a = window.webExcel, d = a.docs.find((x) => x.wb.name.startsWith(n));
      a.activateDoc(a.docs.indexOf(d));
      return d.wb.sheet.charts.map((m) => { const c = a.charts.config(m, d.wb.sheet); return { title: m.title, type: c.type, colors: c.data.datasets.map((ds) => ds.backgroundColor), ds: c.data.datasets.map((ds) => ({ pr: ds.pointRadius, bw: ds.borderWidth, cat: ds.categoryPercentage, bar: ds.barPercentage })), titleSize: c.options.plugins.title.font.size, titleColor: c.options.plugins.title.color, tick: c.options.scales?.y?.ticks?.color, grid: c.options.scales?.y?.grid?.color }; });
    }, name);

    let c = await cfgs('fresh');
    const bar = c.find((x) => x.title === '막대'), line = c.find((x) => x.type === 'line');
    check('new workbook: series colors are the Office 2013+ accents', JSON.stringify(bar.colors) === JSON.stringify(['#4472C4', '#ED7D31', '#A5A5A5']), JSON.stringify(bar.colors));
    check('title 14 pt gray, ticks gray, gridlines light (same numbers as the saved file)', Math.abs(bar.titleSize - 14 * 4 / 3) < 0.01 && bar.titleColor === '#595959' && bar.tick === '#595959' && bar.grid === '#D9D9D9', JSON.stringify(bar));
    check('columns: Excel gap 219 / overlap -27 as bar spacing', Math.abs(bar.ds[0].bar - 1 / 1.27) < 0.001 && Math.abs(bar.ds[0].cat - (3 * 1.27) / (3 * 1.27 + 2.19)) < 0.001, JSON.stringify(bar.ds[0]));
    check('line: 2.25 pt, no markers', line.ds[0].pr === 0 && Math.abs(line.ds[0].bw - 3) < 0.01, JSON.stringify(line.ds[0]));

    // combo chart: B as a line between two bars keeps the color of its position
    const combo = await p.evaluate(() => {
      const a = window.webExcel, d = a.docs.find((x) => x.wb.name.startsWith('fresh'));
      const m = { type: 'bar', series: [{ val: 'Sheet1!$B$2:$B$5', name: 'A' }, { val: 'Sheet1!$C$2:$C$5', name: 'B', type: 'line' }, { val: 'Sheet1!$D$2:$D$5', name: 'C' }] };
      return a.charts.config(m, d.wb.sheet).data.datasets.map((ds) => ds.backgroundColor);
    });
    check('combo: preview colors by series position (= saved file)', JSON.stringify(combo) === JSON.stringify(['#4472C4', '#ED7D31', '#A5A5A5']), JSON.stringify(combo));
    const onePie = await p.evaluate(() => {
      const a = window.webExcel, d = a.docs.find((x) => x.wb.name.startsWith('fresh'));
      const m = { type: 'pie', varyColors: false, series: [{ val: 'Sheet1!$B$2:$B$5', color: '#C00000' }] };
      return a.charts.config(m, d.wb.sheet).data.datasets[0].backgroundColor;
    });
    check('one-color pie: preview draws every slice in its color', onePie.every((x) => x === '#C00000'), JSON.stringify(onePie));
    c = await cfgs('old');
    const colored = c.find((x) => x.title === 'colored'), slices = c.find((x) => x.title === 'slices');
    check('chart with its own colors: first series C00000, second from the workbook theme (Office 2007 accent2)', JSON.stringify(colored.colors) === JSON.stringify(['#C00000', '#C0504D']), JSON.stringify(colored.colors));
    check('pie: its own slice color, then the theme', slices.colors[0][0] === '#112233' && slices.colors[0][1] === '#C0504D', JSON.stringify(slices.colors));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
