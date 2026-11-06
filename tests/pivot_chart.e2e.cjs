// Pivot charts in files (step 1: keep and show). A chart whose DrawingML has c:pivotSource is linked
// to a pivot table: it is listed as a pivot chart, keeps that link when it is edited and saved, and
// refuses what Excel refuses on a pivot chart (changing its data range, switching row/column,
// scatter type). Starts its own server (port 8876).   node tests/pivot_chart.e2e.cjs
// Optional: WX_CHART_XSD=<ECMA-376 transitional dml-chart.xsd> validates the saved chart part.
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-pchart-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const xlErr = (...a) => { try { xl(...a); return null; } catch (e) { return { code: e.status, out: String(e.stderr || e.stdout) }; } };
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });

// the pivot sample (a real pivot written by LibreOffice) plus a pivot chart on its pivot 피벗1 (A5:D10)
py(`
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.pivot import PivotSource, PivotFormat
wb = openpyxl.load_workbook(${JSON.stringify(path.join(ROOT, 'samples', 'pivot.xlsx'))}); ws = wb['피벗']
c = BarChart(); c.title = 'by region'
c.add_data(Reference(ws, min_col=2, max_col=3, min_row=6, max_row=9), titles_from_data=True)
c.set_categories(Reference(ws, min_col=1, min_row=7, max_row=9))
c.pivotSource = PivotSource(name='[pc.xlsx]피벗!피벗1', fmtId=0)
c.pivotFormats = (PivotFormat(idx=0),)
ws.add_chart(c, 'F5')
wb.save(${JSON.stringify(path.join(ws, 'pc.xlsx'))})
`);
const chartXml = (file) => JSON.parse(py(`
import json, re, zipfile
z = zipfile.ZipFile(${JSON.stringify(path.join(ws, '__F__'))}.replace('__F__', ${JSON.stringify(file)}))
out = {}
for n in z.namelist():
    if n.startswith('xl/charts/chart'):
        s = z.read(n).decode()
        out[n] = {'pivotSource': re.findall(r'<pivotSource>.*?</pivotSource>', s, re.S), 'pivotFmts': '<pivotFmts>' in s,
                  'type': re.findall(r'<(barChart|lineChart|scatterChart)>', s), 'title': re.findall(r'<a:t>([^<]*)', s)[:1], 'f': re.findall(r'<f>([^<]*)</f>', s)}
print(json.dumps(out, ensure_ascii=False))`));

(async () => {
  xl('serve', '--port', '8876'); xl('open', 'pc.xlsx');
  const ls = xl('chart', 'ls');
  check('chart ls: listed as a pivot chart of its pivot table', /'by region'\s+@F5\s+pivot chart of '피벗1'/.test(ls), ls);
  const d = JSON.parse(xl('describe', '--json'));
  const ch = d.sheets.find((s) => s.name === '피벗').charts[0];
  check('describe: chart carries pivot: 피벗1', ch.pivot === '피벗1', JSON.stringify(ch));
  const id = ch.id;
  // what Excel refuses on a pivot chart
  let e = xlErr('chart', 'set', id, '--data', '피벗!A6:C9');
  check('changing its data range is refused (says to change the pivot table)', e && e.code !== 0 && /pivot chart of pivot table '피벗1'/.test(e.out), JSON.stringify(e));
  e = xlErr('chart', 'set', id, '--switch');
  check('switching row/column is refused', e && e.code !== 0 && /pivot chart/.test(e.out), JSON.stringify(e));
  e = xlErr('chart', 'set', id, '--to-type', 'scatter');
  check('scatter type is refused (Excel: no XY scatter pivot charts)', e && e.code !== 0 && /cannot be a scatter chart/.test(e.out), JSON.stringify(e));
  // untouched, then edited: the link stays
  xl('save', '--as', 'untouched.xlsx');
  let x = Object.values(chartXml('untouched.xlsx'))[0];
  check('saved untouched: pivotSource kept', x.pivotSource[0]?.includes('[pc.xlsx]피벗!피벗1'), JSON.stringify(x));
  xl('chart', 'set', id, '--title', 'by region (edited)', '--to-type', 'line');
  xl('save', '--as', 'edited.xlsx');
  x = Object.values(chartXml('edited.xlsx'))[0];
  check('saved after edits (title, type line): still a pivot chart, same link and fmtId, pivotFmts kept',
    x.pivotSource[0]?.includes('[pc.xlsx]피벗!피벗1') && /fmtId val="0"/.test(x.pivotSource[0]) && x.pivotFmts && x.type.join() === 'lineChart' && x.title[0] === 'by region (edited)', JSON.stringify(x));
  if (process.env.WX_CHART_XSD) {
    const ok = py(`
import zipfile
from lxml import etree
sc = etree.XMLSchema(etree.parse(${JSON.stringify(process.env.WX_CHART_XSD)}))
z = zipfile.ZipFile(${JSON.stringify(path.join(ws, 'edited.xlsx'))})
print(all(sc.validate(etree.fromstring(z.read(n))) for n in z.namelist() if n.startswith('xl/charts/chart')))`).trim();
    check('…the chart part is valid against ECMA-376 dml-chart.xsd', ok === 'True', ok);
  } else console.log('SKIP XSD check (set WX_CHART_XSD to the ECMA-376 transitional dml-chart.xsd)');

  // preview
  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 800 } });
    p.on('pageerror', (er) => errors.push(er.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((s) => s.name === '피벗')); });
    await sleep(500);
    const cid = await p.evaluate(() => window.monoonom.doc.wb.sheet.charts[0].id);
    await p.evaluate((i) => window.monoonom.charts.select(i), cid); await sleep(250);
    const panel = await p.evaluate(() => ({ text: document.querySelector('.chart-ranges')?.textContent || '', sw: !!document.querySelector('.chart-ranges [data-switch]') }));
    check('preview: the range panel says it is a pivot chart of 피벗1 and has no switch button', /피벗 테이블 '피벗1'의 피벗 차트/.test(panel.text) && !panel.sw, JSON.stringify(panel));
    await p.evaluate((i) => window.monoonom.charts.switchChart(i), cid); await sleep(300);
    const toast = await p.evaluate(() => [...document.querySelectorAll('.toast')].map((x) => x.textContent).join(' | '));
    check('…switching row/column from the menu is refused with a message', /피벗 차트/.test(toast), toast);
    await p.evaluate(() => { window.monoonom.chartDialog(window.monoonom.doc.wb.sheet.charts[0]); }); await sleep(300);
    const dlg = await p.evaluate(() => ({ range: document.querySelector('.chart-dialog [data-k=range]')?.disabled, scatter: document.querySelector('.chart-dialog [data-t=scatter]')?.disabled }));
    check('…edit dialog: data range and scatter are disabled', dlg.range === true && dlg.scatter === true, JSON.stringify(dlg));
    await p.fill('.chart-dialog [data-k=title]', 'from the preview');
    await p.click('.chart-dialog [data-act=ok]'); await sleep(600);
    // step 2: the pivot chart follows its pivot table
    const data = () => p.evaluate(() => { const a = window.monoonom, m = a.doc.wb.sheet.charts[0]; return a.charts.instances.get(m.id).chart.data.datasets.map((d) => d.data.join(',')); });
    const before = await data();
    xl('write', '데이터!D2', '1000');  // 서울 A Q1 120 -> 1000
    const ref = xl('pivot', 'refresh');
    check('refresh: names the pivot charts that show the new values', ref.includes(`pivot charts ${cid} show them`), ref);
    await p.waitForFunction((b) => { const a = window.monoonom, m = a.doc.wb.sheet.charts[0]; return a.charts.instances.get(m.id)?.chart.data.datasets[0].data.join(',') !== b; }, before[0], { timeout: 4000 }).catch(() => {});
    const after = await data();
    check('…the preview pivot chart redraws with the refreshed values (서울 A 250 -> 1130)', before[0] === '85,125,250' && after[0] === '85,125,1130', `${before[0]} -> ${after[0]}`);
    xl('rows', 'insert', '2', '--count', '2', '--sheet', '피벗');
    const moved = xl('chart', 'ls');
    check('inserting rows above the pivot moves it, and the pivot chart\'s ranges move with it', /A7:D12/.test(xl('pivot', 'ls')) && /Y '?피벗'?!\$B\$9:\$B\$11/.test(moved), moved);
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  xl('save', '--as', 'preview.xlsx');
  x = Object.values(chartXml('preview.xlsx'))[0];
  check('saved after editing in the preview, refresh and row insert: still linked to 피벗1, ranges moved', x.pivotSource[0]?.includes('피벗1') && x.title[0] === 'from the preview' && x.f.some((f) => /^'?피벗'?!\$B\$9:\$B\$11$/.test(f)), JSON.stringify(x));
  try { xl('stop', '--force'); } catch { /* ignore */ }
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((er) => { console.error(er); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
