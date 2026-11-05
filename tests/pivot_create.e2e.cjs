// Pivot tables, step 3: create one (agent `pivot add` and the preview's dialog), refresh it, save it
// as a real pivot table (ECMA-376 XSD-valid; LibreOffice reads it and computes the same numbers).
// Starts its own server (port 8866).   node tests/pivot_create.e2e.cjs
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'web-excel', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-pivotc-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'pivot.xlsx'), path.join(ws, 'pivot.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const xlErr = (...a) => { try { xl(...a); return null; } catch (e) { return { code: e.status, out: String(e.stderr || e.stdout) }; } };
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });
const XSD = process.env.WX_SML_XSD; // optional: ECMA-376 Part 4 transitional sml.xsd for the schema check

(async () => {
  xl('serve', '--port', '8866'); xl('open', 'pivot.xlsx');
  // agent: two row fields (subtotals), two data fields, on a new sheet
  const out = xl('pivot', 'add', '데이터!A1', '--rows', '지역,분기', '--values', '매출:sum,매출:average', '--at', '요약!B2');
  check('agent pivot add: created on a new sheet', /피벗 테이블 '피벗 테이블1' → 요약!B2:E13/.test(out), out);
  const rd = xl('read', '요약!B2:E13');
  check('…tabular layout with a 값 header, subtotals and grand total', /값/.test(rd) && /대구 요약\t\t150\t37\.5/.test(rd) && /총합계\t\t860\t71\.66/.test(rd), rd);
  let e = xlErr('pivot', 'add', '데이터!A1', '--rows', '지역', '--values', '매출', '--at', '데이터!B2');
  check('agent: placing it over data is refused', e && e.code === 4 && /데이터가 있습니다/.test(e.out), JSON.stringify(e));
  e = xlErr('pivot', 'add', '데이터!A1', '--rows', '도시', '--values', '매출', '--at', '새!A1');
  check('agent: an unknown field is refused with the header names', e && /'도시' 열이 없습니다. 머리글: 지역, 제품, 분기, 매출/.test(e.out), JSON.stringify(e));
  e = xlErr('pivot', 'add', '데이터!A1', '--rows', '지역', '--cols', '제품', '--values', '매출:sum,매출:max', '--at', '새!A1');
  check('agent: columns + several data fields is refused (not supported yet)', e && /값 필드는 하나까지/.test(e.out), JSON.stringify(e));
  check('agent: the new pivot is protected like a file pivot', /영역은 바꿀 수 없습니다/.test(JSON.stringify(xlErr('write', '요약!D5', '1'))));

  const browser = await chromium.launch();
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 800 } });
    p.on('pageerror', (er) => errors.push(er.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.webExcel?.doc?.wb.sheets.some((s) => s.name === '요약')); await sleep(300);
    // the dialog: rows 제품, columns 지역, values 매출 (sum), filter 분기, new sheet
    await p.evaluate(() => { const a = window.webExcel; a.activateSheet(0); a.select(0, 0); a.exec('pivotDialog'); });
    await p.waitForSelector('.pivot-dialog .pv-row');
    const set = async (f, role) => p.selectOption(`.pivot-dialog .pv-row[data-f="${f}"] [data-role]`, role);
    await set('지역', 'cols'); await set('제품', 'rows'); await set('분기', 'filters'); await set('매출', 'values');
    await p.fill('.pivot-dialog [data-k=sheet]', '제품별');
    await p.click('.pivot-dialog .modal-btns .primary');
    await p.waitForFunction(() => window.webExcel.doc.wb.sheet.name === '제품별', null, { timeout: 6000 }).catch(() => {});
    const ui = await p.evaluate(() => { const s = window.webExcel.doc.wb.sheet; const v = (r, c) => s.get(r, c)?.v ?? null;
      return { sheet: s.name, pivot: s.pivots?.[0]?.ref, filter: [v(2, 0), v(2, 1)], head: [v(5, 0), v(5, 1), v(5, 4)], a: [v(6, 1), v(6, 2), v(6, 3), v(6, 4)], grand: v(8, 4) }; });
    check('dialog: pivot on a new sheet, shown; filter row (분기 / (모두)) above it', ui.sheet === '제품별' && ui.pivot === 'A5:E9' && ui.filter.join() === '분기,(모두)', JSON.stringify(ui));
    check('…product A row: 대구 85, 부산 125, 서울 250, total 460; grand 860', ui.a.join() === '85,125,250,460' && ui.grand === 860 && ui.head.join() === '제품,대구,총합계', JSON.stringify(ui));
    check('…closes the dialog', !(await p.$('.pivot-dialog')));
    if (process.env.SHOT) await p.screenshot({ path: process.env.SHOT });
    // refresh the created pivots after a source change
    xl('write', '데이터!D2', '1000');
    const ref = JSON.parse(xl('pivot', 'refresh', '--json'));
    check('refresh: the created pivots update (서울 A Q1 120 -> 1000)', ref.pivots.filter((x) => x.changed > 0).map((x) => x.name).sort().join() === '피벗 테이블1,피벗 테이블2,피벗1', JSON.stringify(ref.pivots));
    await p.waitForFunction(() => window.webExcel.doc.wb.sheets.find((s) => s.name === '제품별').get(6, 3)?.v === 1130, null, { timeout: 4000 }).catch(() => {});
    check('…and the preview shows it (서울 A 1130)', await p.evaluate(() => window.webExcel.doc.wb.sheets.find((s) => s.name === '제품별').get(6, 3)?.v === 1130));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  // save twice: real pivot tables, no duplicates
  xl('save', '--as', 'out.xlsx'); xl('write', '데이터!F1', 'x'); xl('save');
  const r = JSON.parse(py(`
import json, openpyxl, zipfile, re
f = ${JSON.stringify(path.join(ws, 'out.xlsx'))}
wb = openpyxl.load_workbook(f)
res = {ws.title: [(p.name, p.location.ref, bool(p.cache and p.cache.refreshOnLoad)) for p in ws._pivots] for ws in wb}
z = zipfile.ZipFile(f)
res['_parts'] = sorted(n for n in z.namelist() if 'pivot' in n and n.endswith('.xml') and '_rels' not in n)
res['_rid'] = any(re.search(r'<pivotTableDefinition[^>]* r:id=', z.read(n).decode()) for n in res['_parts'] if 'pivotTables' in n)
xsd = ${JSON.stringify(XSD || '')}
if xsd:
    from lxml import etree
    sc = etree.XMLSchema(etree.parse(xsd))
    res['_xsd'] = {n: sc.validate(etree.fromstring(z.read(n))) for n in res['_parts'] + ['xl/workbook.xml']}
print(json.dumps(res, ensure_ascii=False))`));
  check('saved twice: 3 pivots, no duplicates; new ones refresh on open', JSON.stringify(r['요약']) === '[["피벗 테이블1","B2:E13",true]]'
    && r['제품별'].length === 1 && r['제품별'][0][1] === 'A5:E9' && r['피벗'].length === 1, JSON.stringify(r));
  check('…no r:id on pivotTableDefinition (not in the schema)', r._rid === false);
  if (XSD) check('…every pivot part and workbook.xml valid against ECMA-376 sml.xsd', Object.values(r._xsd).every(Boolean), JSON.stringify(r._xsd));
  else console.log('SKIP XSD check (set WX_SML_XSD to the ECMA-376 transitional sml.xsd)');
  // LibreOffice reads the file as pivot tables and computes the same numbers (when installed)
  let lo = null;
  try { lo = execFileSync('which', ['soffice'], { encoding: 'utf8' }).trim(); } catch {}
  if (lo) {
    const od = path.join(ws, 'lo'); fs.mkdirSync(od);
    execFileSync(lo, ['--headless', '--convert-to', 'xlsx', '--outdir', od, path.join(ws, 'out.xlsx')], { stdio: 'ignore', timeout: 180000 });
    const cmp = JSON.parse(py(`
import json, openpyxl
from openpyxl.utils.cell import range_boundaries
a = openpyxl.load_workbook(${JSON.stringify(path.join(ws, 'out.xlsx'))}); b = openpyxl.load_workbook(${JSON.stringify(path.join(ws, 'lo', 'out.xlsx'))})
def nums(ws, ref):
    c1, r1, c2, r2 = range_boundaries(ref)
    return [[round(v, 6) for v in (ws.cell(r, c).value for c in range(c1, c2 + 1)) if isinstance(v, (int, float))] for r in range(r1, r2 + 1)]
out = {}
for sh in ('요약', '제품별'):
    pa, pb = a[sh]._pivots[0], b[sh]._pivots[0]
    out[sh] = {'loc': [pa.location.ref, pb.location.ref], 'same': nums(a[sh], pa.location.ref) == nums(b[sh], pa.location.ref)}
print(json.dumps(out, ensure_ascii=False))`));
    check('LibreOffice: reads both created pivots at the same place and computes the same numbers',
      Object.values(cmp).every((x) => x.loc[0] === x.loc[1] && x.same), JSON.stringify(cmp));
  } else console.log('SKIP LibreOffice check (soffice not installed)');
  try { xl('stop', '--force'); } catch {}
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((er) => { console.error(er); try { xl('stop', '--force'); } catch {} process.exit(1); });
