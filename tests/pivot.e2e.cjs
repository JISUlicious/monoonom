// Pivot tables (step 1: safeguards) in the preview and for the agent, with samples/pivot.xlsx
// (a real pivot written by LibreOffice: tests/make_pivot_sample.py). Starts its own server (port 8863).
//   node tests/pivot.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-pivot-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'pivot.xlsx'), path.join(ws, 'pivot.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const xlErr = (...a) => { try { xl(...a); return null; } catch (e) { return { code: e.status, out: String(e.stderr || e.stdout) }; } };
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });

(async () => {
  xl('serve', '--port', '8863');
  const opened = xl('open', 'pivot.xlsx');
  check('open: the pivot is reported (kept, not editable)', /Pivot tables on sheet '피벗' \(피벗1\) are kept and their areas cannot be edited/.test(opened), opened);
  const d = JSON.parse(xl('describe', '--json'));
  const sh = d.sheets.find((s) => s.name === '피벗');
  check('describe: pivot with range, source and fields', sh.pivots.length === 1 && sh.pivots[0].range === 'A5:D10' && sh.pivots[0].source === '데이터!A1:D13'
    && sh.pivots[0].rows.join() === '지역' && sh.pivots[0].cols.join() === '제품' && sh.pivots[0].values.join() === 'sum(매출)', JSON.stringify(sh.pivots));
  check('describe: the pivot output is not listed as a table', !sh.tables.some((t) => /^A5/.test(t.range)), JSON.stringify(sh.tables));
  // agent side
  let e = xlErr('write', '피벗!B7', '1');
  check('agent: writing into the pivot is refused (exit 4, says to edit the source)', e && e.code === 4 && /area of pivot table '피벗1' \(A5:D10\) cannot be changed/.test(e.out), JSON.stringify(e));
  check('agent: formatting the pivot is allowed', /format/.test(xl('format', '피벗!A5:D5', 'b=true')));
  e = xlErr('rows', 'insert', '7', '--sheet', '피벗');
  check('agent: inserting rows inside the pivot is refused', e && e.code === 4 && /Cannot insert rows inside pivot table/.test(e.out), JSON.stringify(e));
  e = xlErr('sort', '피벗!A1:D12', '--by', 'A');
  check('agent: sorting a range with the pivot is refused', e && /Cannot sort a range/.test(e.out), JSON.stringify(e));

  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 760 } });
    p.on('pageerror', (er) => errors.push(er.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((s) => s.name === '피벗')); });
    await sleep(300);
    check('preview: the sheet knows its pivot', await p.evaluate(() => window.monoonom.doc.wb.sheet.pivots?.[0]?.ref === 'A5:D10'));
    const toasts = () => p.evaluate(() => [...document.querySelectorAll('.toast')].map((t) => t.textContent));
    const val = (r, c) => p.evaluate(([r, c]) => window.monoonom.doc.wb.sheets.find((s) => s.name === '피벗').get(r, c)?.v ?? null, [r, c]);
    // typing into a pivot cell
    const before = await val(6, 1);
    await p.evaluate(() => { const a = window.monoonom; a.select(6, 1); a.focusGrid(); });
    await p.keyboard.type('777'); await p.keyboard.press('Enter'); await sleep(300);
    check('preview: typing into the pivot is undone and explained', (await val(6, 1)) === before && (await toasts()).some((t) => /피벗 테이블 '피벗1'/.test(t)), `${await val(6, 1)} ${JSON.stringify(await toasts())}`);
    // Delete key
    await p.evaluate(() => { const a = window.monoonom; a.selectRange({ r1: 6, c1: 1, r2: 8, c2: 2 }); a.focusGrid(); });
    await p.keyboard.press('Delete'); await sleep(300);
    check('preview: Delete on pivot cells keeps them', (await val(6, 1)) === before && (await val(8, 2)) !== null);
    // a cell outside still works
    await p.evaluate(() => { const a = window.monoonom; a.select(13, 0); a.focusGrid(); });
    await p.keyboard.type('밖'); await p.keyboard.press('Enter'); await sleep(300);
    check('preview: cells outside the pivot are editable', (await val(13, 0)) === '밖');
    // rows: inside refused, above moves the pivot
    const inside = await p.evaluate(() => { const a = window.monoonom; a.select(7, 0); a.exec('insertRow'); return window.monoonom.doc.wb.sheet.pivots[0].ref; });
    await sleep(200);
    check('preview: inserting a row inside the pivot is refused', inside === 'A5:D10' && (await toasts()).some((t) => /가운데에는 행을 삽입할 수 없습니다/.test(t)), inside);
    await p.evaluate(() => { const a = window.monoonom; a.select(1, 0); a.exec('insertRow'); });
    await sleep(600);
    const moved = await p.evaluate(() => window.monoonom.doc.wb.sheet.pivots[0].ref);
    const srv = JSON.parse(xl('describe', '--json')).sheets.find((s) => s.name === '피벗').pivots[0].range;
    check('preview: a row above moves the pivot (A6:D11), the server agrees', moved === 'A6:D11' && srv === 'A6:D11', `${moved} / ${srv}`);
    check('…its cells moved with it', (await val(5, 0)) === 'Sum - 매출', String(await val(5, 0)));
    if (process.env.SHOT) await p.screenshot({ path: process.env.SHOT });
    // the server refuses pivot writes even without the preview's check (another client)
    const raw = await p.evaluate(async () => {
      const a = window.monoonom, d = a.doc;
      const r = await fetch(`/api/docs/${d.wb.docId}/ops`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Monoonom': '1' },
        body: JSON.stringify({ ops: [{ op: 'setCells', sheet: d.wb.sheet.id, cells: [[6, 1, { v: 1 }]] }], baseVersion: d.sync.version }) });
      return { status: r.status, body: await r.text() };
    });
    check('server: a raw op into the pivot is refused (409)', raw.status === 409 && raw.body.includes('피벗'), JSON.stringify(raw).slice(0, 200));
    // step 2: refresh in place — the agent changes the source, the user refreshes from the context menu
    xl('write', '데이터!D2', '1000');  // 서울 A Q1: 120 -> 1000
    await sleep(500);
    check('before refresh the pivot still shows the old value (서울 A 250)', (await val(9, 1)) === 250, String(await val(9, 1)));
    // (follow mode moved the view to the agent's sheet): back to the pivot
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((s) => s.name === '피벗')); a.select(9, 1); });
    await sleep(200);
    const box = await p.evaluate(() => { const g = window.monoonom.grid, m = g.metrics(); return { x: g.xOf(1, m) + 20, y: g.yOf(9, m) + 8, rect: g.canvas.getBoundingClientRect().toJSON() }; }).catch(() => null);
    let usedMenu = false;
    if (box) {
      await p.mouse.click(box.rect.x + box.x, box.rect.y + box.y, { button: 'right' }); await sleep(200);
      const item = p.locator('.ctx-menu button', { hasText: '피벗 새로 고침' }).first();
      if (await item.count()) { await item.click(); usedMenu = true; }
    }
    if (!usedMenu) await p.evaluate(() => window.monoonom.exec('pivotRefresh'));
    await p.waitForFunction(() => window.monoonom.doc.wb.sheets.find((s) => s.name === '피벗').get(9, 1)?.v === 1130, null, { timeout: 5000 }).catch(() => {});
    check(`preview refresh${usedMenu ? ' (context menu)' : ' (command)'}: 서울 A 1130, 서울 total 1280, grand 1740`,
      (await val(9, 1)) === 1130 && (await val(9, 3)) === 1280 && (await val(10, 3)) === 1740, `${await val(9, 1)} ${await val(9, 3)} ${await val(10, 3)}`);
    check('…the toast says how many values changed', (await toasts()).some((t) => /피벗 새로 고침: 피벗1 — 값 4개 바뀜/.test(t)), JSON.stringify(await toasts()));
    const agentSees = xl('read', '피벗!B10:D11');
    check('…the agent reads the refreshed values', /1130/.test(agentSees) && /1740/.test(agentSees), agentSees);
    const ref = JSON.parse(xl('pivot', 'refresh', '--json'));
    check('agent `pivot refresh` right after: nothing left to change', ref.pivots[0].changed === 0 && !ref.batchId, JSON.stringify(ref));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  // the source changed (above), save: Excel is asked to refresh; location kept in step
  xl('save', '--as', 'out.xlsx');
  const r = JSON.parse(py(`
import json, openpyxl
wb = openpyxl.load_workbook(${JSON.stringify(path.join(ws, 'out.xlsx'))}); p = wb['피벗']._pivots[0]
print(json.dumps({'loc': p.location.ref, 'refresh': p.cache.refreshOnLoad, 'src': p.cache.cacheSource.worksheetSource.ref, 'top': wb['피벗']['A6'].value}))`));
  check('saved: pivot at A6:D11 with its cells, source unchanged, refreshOnLoad on (source data changed)', r.loc === 'A6:D11' && r.top === 'Sum - 매출' && r.src === 'A1:D13' && r.refresh === true, JSON.stringify(r));
  try { xl('stop', '--force'); } catch {}
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((er) => { console.error(er); try { xl('stop', '--force'); } catch {} process.exit(1); });
