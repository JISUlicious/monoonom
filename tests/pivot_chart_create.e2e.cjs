// Pivot charts, step 3: make one from a pivot table (agent: chart add --pivot NAME; preview: Insert →
// PivotChart, or the pivot's context menu). Series are the pivot's data columns and categories its row
// items, totals left out; the chart is linked by c:pivotSource "[<saved file>]Sheet!Pivot".
// Starts its own server (port 8878).   node tests/pivot_chart_create.e2e.cjs
// Optional: WX_CHART_XSD=<ECMA-376 transitional dml-chart.xsd> validates the saved chart parts.
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-pchartc-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'pivot.xlsx'), path.join(ws, 'pivot.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const xlErr = (...a) => { try { xl(...a); return null; } catch (e) { return { code: e.status, out: String(e.stderr || e.stdout) }; } };
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });
const chartsOf = (ls, sheet) => ls.split('\n').filter((l) => l.startsWith(sheet + '  ch'));

(async () => {
  xl('serve', '--port', '8878'); xl('open', 'pivot.xlsx');
  // agent: the LibreOffice pivot 피벗1 (A5:D10: 지역 x 제품, "Total Result" row and column)
  let out = xl('chart', 'add', '--pivot', '피벗1', '--title', 'by region');
  check('agent: chart add --pivot makes a pivot chart next to the pivot', /add pivot chart of '피벗1' \(bar/.test(out), out);
  let ls = xl('chart', 'ls');
  check('…series = the item columns A, B (names in row 6), categories = the regions (rows 7-9); totals left out',
    /@F5\s+pivot chart of '피벗1'/.test(ls) && /series 1: name '?피벗'?!\$B\$6\s+X '?피벗'?!\$A\$7:\$A\$9\s+Y '?피벗'?!\$B\$7:\$B\$9/.test(ls)
    && /series 2: name '?피벗'?!\$C\$6/.test(ls) && !/series 3/.test(ls), ls);
  // a pivot made here (item records): one row field and a column field, as a line chart
  xl('pivot', 'add', '데이터!A1', '--rows', '지역', '--cols', '제품', '--values', '매출', '--at', 'P2!A1');
  out = xl('chart', 'add', '--pivot', 'PivotTable1', '--type', 'line', '--at', 'P2!G2');
  ls = xl('chart', 'ls');
  check('agent: a pivot made here gets a line pivot chart where --at says', /add pivot chart of 'PivotTable1' \(line/.test(out) && chartsOf(ls, 'P2').some((l) => /@G2\s+pivot chart of 'PivotTable1'/.test(l)), ls);
  // refused
  xl('pivot', 'add', '데이터!A1', '--rows', '지역,분기', '--values', '매출', '--at', 'P3!A1');
  let e = xlErr('chart', 'add', '--pivot', 'PivotTable2');
  check('two row fields (subtotals): refused with the reason', e && /one row field/.test(e.out), JSON.stringify(e));
  e = xlErr('chart', 'add', '--pivot', '피벗1', '--type', 'scatter');
  check('scatter: refused (Excel has no XY scatter pivot charts)', e && /cannot be a scatter chart/.test(e.out), JSON.stringify(e));
  e = xlErr('chart', 'add', '--pivot', 'Nope');
  check('unknown pivot: not found', e && /Pivot table not found: Nope/.test(e.out), JSON.stringify(e));

  // preview: Insert → PivotChart on a selected pivot cell
  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 800 } });
    p.on('pageerror', (er) => errors.push(er.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    const ribbon = await p.evaluate(() => document.querySelector('[data-cmd=pivotChart]')?.textContent.trim());
    check('preview: Insert tab has a PivotChart button (Korean UI)', ribbon === '📊피벗 차트' || /피벗 차트/.test(ribbon || ''), ribbon);
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((s) => s.name === '데이터')); a.select(1, 1); });
    await p.evaluate(() => window.monoonom.exec('pivotChart')); await sleep(400);
    const warn = await p.evaluate(() => [...document.querySelectorAll('.toast')].map((x) => x.textContent).join(' | '));
    check('…outside a pivot (several pivots): asks to select a pivot cell', /피벗 테이블의 셀을 선택/.test(warn), warn);
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((s) => s.name === 'P2')); a.select(2, 1); });
    const before = await p.evaluate(() => window.monoonom.doc.wb.sheet.charts.length);
    await p.evaluate(() => window.monoonom.exec('pivotChart'));
    await p.waitForFunction((n) => window.monoonom.doc.wb.sheet.charts.length > n, before, { timeout: 4000 }).catch(() => {});
    const made = await p.evaluate(() => { const ch = window.monoonom.doc.wb.sheet.charts.at(-1); return { n: window.monoonom.doc.wb.sheet.charts.length, pivot: ch?.pivot?.table, series: ch?.series?.length }; });
    check('…on a pivot cell: a pivot chart of that pivot appears in the preview', made.n === before + 1 && made.pivot === 'PivotTable1' && made.series === 2, JSON.stringify(made));
    await sleep(400);
    const drawn = await p.evaluate(() => { const a = window.monoonom, m = a.doc.wb.sheet.charts.at(-1); return a.charts.instances.get(m.id)?.chart.data.labels?.join(','); });
    check('…drawn with the region labels', drawn === '대구,부산,서울', drawn);
    const status = await p.evaluate(() => document.getElementById('status-msg').textContent);
    check('…status bar confirms it (Korean)', /피벗 차트를 삽입했습니다/.test(status), status);
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }

  // saved: real pivot charts, named after the saved file
  xl('save', '--as', 'out.xlsx');
  const r = JSON.parse(py(`
import json, re, zipfile
z = zipfile.ZipFile(${JSON.stringify(path.join(ws, 'out.xlsx'))})
out = []
for n in sorted(x for x in z.namelist() if x.startswith('xl/charts/chart')):
    s = z.read(n).decode()
    out.append(re.findall(r'<pivotSource><name>([^<]*)</name>', s))
xsd = ${JSON.stringify(process.env.WX_CHART_XSD || '')}
ok = None
if xsd:
    from lxml import etree
    sc = etree.XMLSchema(etree.parse(xsd))
    ok = all(sc.validate(etree.fromstring(z.read(n))) for n in z.namelist() if n.startswith('xl/charts/chart'))
print(json.dumps({'names': out, 'xsd': ok}, ensure_ascii=False))`));
  const names = r.names.flat().sort();
  check('saved: three pivot charts, each pivotSource named [out.xlsx]<sheet>!<pivot>',
    names.length === 3 && names.includes("[out.xlsx]'P2'!PivotTable1") && names.includes('[out.xlsx]피벗!피벗1'), JSON.stringify(r.names));
  if (process.env.WX_CHART_XSD) check('…every chart part valid against ECMA-376 dml-chart.xsd', r.xsd === true, String(r.xsd));
  else console.log('SKIP XSD check (set WX_CHART_XSD to the ECMA-376 transitional dml-chart.xsd)');
  try { xl('stop', '--force'); } catch { /* ignore */ }
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((er) => { console.error(er); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
