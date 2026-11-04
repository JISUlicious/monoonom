// P1 sync test: server-owned documents shared by the browser and an "agent" (HTTP client).
//   python3 run.py --port 8765 --root . &   node tests/sync.e2e.cjs
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const state = JSON.parse(fs.readFileSync(path.join(ROOT, '.web-excel', 'server.json'), 'utf8'));
const BASE = `http://127.0.0.1:${state.port}`;
let failures = 0;
const check = (name, cond, extra = '') => { console.log(`${cond ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!cond) failures++; };
const agent = async (method, p, body) => {
  const res = await fetch(BASE + p, { method, headers: { 'X-WebExcel-Token': state.token, 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  const j = await res.json().catch(() => ({}));
  return { status: res.status, ...j };
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1300, height: 800 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(BASE + '/');
  await page.waitForFunction(() => window.webExcel?.started);
  await page.evaluate(async () => { if (!window.webExcel.doc) await window.webExcel.newDoc(); }); // no workbook is created automatically
  await page.waitForFunction(() => window.webExcel?.doc?.sync);
  const docId = await page.evaluate(() => window.webExcel.doc.wb.docId);
  const sid = await page.evaluate(() => window.webExcel.doc.wb.sheet.id);
  const val = (r, c) => page.evaluate(([r, c]) => { const a = window.webExcel; const v = a.doc.wb.value(a.doc.wb.sheet, r, c); return v && v.err ? v.err : v; }, [r, c]);

  // 1) agent writes -> browser shows it (and highlights)
  let r = await agent('POST', `/api/docs/${docId}/ops`, { actor: 'agent', clientId: 'cli', label: '값 입력', ops: [{ op: 'setCells', sheet: sid, cells: [[0, 0, { v: 7 }], [0, 1, { f: '=A1*6' }]] }] });
  check('agent ops accepted', r.status === 200 && r.version >= 1, JSON.stringify(r));
  await page.waitForFunction(() => window.webExcel.doc.wb.sheet.get(0, 1)?.f === '=A1*6', null, { timeout: 5000 }).catch(() => {});
  check('agent edit appears in browser', (await val(0, 1)) === 42, String(await val(0, 1)));
  check('status shows agent change', /에이전트/.test(await page.evaluate(() => document.querySelector('#status-msg').textContent)));

  // 2) browser edit -> server snapshot
  await page.evaluate(() => window.webExcel.select(2, 0));
  await page.keyboard.type('브라우저'); await page.keyboard.press('Enter');
  await sleep(500);
  let snap = await agent('GET', `/api/docs/${docId}/snapshot`);
  const cellOf = (s, r2, c2) => s.sheets[0].cells.find(([rr, cc]) => rr === r2 && cc === c2)?.[2];
  check('browser edit reaches server', cellOf(snap, 2, 0)?.v === '브라우저', JSON.stringify(cellOf(snap, 2, 0)));

  // 3) browser reports computed formula values (calc source for the agent)
  await sleep(700);
  const doc = await agent('GET', `/api/docs/${docId}`);
  check('calc reported up to current version', doc.calcVersion >= doc.version - 1, `calc=${doc.calcVersion} v=${doc.version}`);

  // 4) concurrent writes to the same cell converge to one value everywhere
  await page.evaluate(() => window.webExcel.select(5, 5));
  const typing = page.keyboard.type('사용자').then(() => page.keyboard.press('Enter'));
  await agent('POST', `/api/docs/${docId}/ops`, { actor: 'agent', clientId: 'cli', ops: [{ op: 'setCells', sheet: sid, cells: [[5, 5, { v: '에이전트' }]] }] });
  await typing; await sleep(800);
  snap = await agent('GET', `/api/docs/${docId}/snapshot`);
  const serverV = cellOf(snap, 5, 5)?.v, browserV = await val(5, 5);
  check('same-cell race converges', serverV === browserV, `server=${serverV} browser=${browserV}`);

  // 5) stale structural batch -> 409; browser rename (structural) after an agent change still works
  r = await agent('POST', `/api/docs/${docId}/ops`, { actor: 'agent', clientId: 'cli2', baseVersion: 0, structural: true, ops: [] });
  check('stale structural batch rejected', r.status === 409, String(r.status));

  // 6) second window sees the same document and live edits
  const page2 = await browser.newPage();
  await page2.goto(`${BASE}/?doc=${docId}`);
  await page2.waitForFunction((id) => window.webExcel?.doc?.wb.docId === id, docId);
  await page2.evaluate(() => window.webExcel.select(8, 0));
  await page2.keyboard.type('창2'); await page2.keyboard.press('Enter');
  await page.waitForFunction(() => window.webExcel.doc.wb.sheet.get(8, 0)?.v === '창2', null, { timeout: 5000 }).catch(() => {});
  check('window 2 edit appears in window 1', (await val(8, 0)) === '창2');

  // 7) agent opens a file -> it appears as a tab in the browser
  const opened = await agent('POST', '/api/docs', { path: path.join(ROOT, 'samples', 'sample.xlsx') });
  await page.waitForFunction((id) => window.webExcel.docs.some((d) => d.wb.docId === id), opened.docId, { timeout: 5000 }).catch(() => {});
  check('agent-opened file appears in browser', await page.evaluate((id) => window.webExcel.docs.some((d) => d.wb.docId === id), opened.docId));

  // 8) agent batch undo
  r = await agent('POST', `/api/docs/${opened.docId}/ops`, { actor: 'agent', clientId: 'cli', batchId: 'b-undo-test', label: '테스트', ops: [{ op: 'setCells', sheet: (await agent('GET', `/api/docs/${opened.docId}/snapshot`)).sheets[0].id, cells: [[0, 0, { v: 'X' }]] }] });
  r = await agent('POST', `/api/docs/${opened.docId}/undo`, { batchId: 'b-undo-test' });
  snap = await agent('GET', `/api/docs/${opened.docId}/snapshot`);
  check('agent batch undo restores value', cellOf(snap, 0, 0)?.v === '월', JSON.stringify(cellOf(snap, 0, 0)));

  // 9) P3: agent row insert appears in the browser with formulas shifted
  await page.evaluate((id) => window.webExcel.activateDoc(window.webExcel.docs.findIndex((d) => d.wb.docId === id)), opened.docId);
  const sid2 = (await agent('GET', `/api/docs/${opened.docId}/snapshot`)).sheets[0].id;
  r = await agent('POST', `/api/docs/${opened.docId}/edit`, { edits: [{ op: 'insertRows', sheet: '매출', at: '3', count: 2 }] });
  await page.waitForFunction(() => window.webExcel.doc.wb.sheet.get(9, 1)?.f === '=SUM(B2:B9)', null, { timeout: 5000 }).catch(() => {});
  check('agent insertRows reflected in browser', await page.evaluate(() => window.webExcel.doc.wb.sheet.get(9, 1)?.f) === '=SUM(B2:B9)');

  // 10) user edit in flight while the agent inserts rows above it -> moved on both sides
  let release;
  const gate = new Promise((res) => { release = res; });
  await page.route('**/api/docs/*/ops', async (route) => { await gate; await route.continue(); });
  await page.evaluate(() => window.webExcel.select(11, 7)); // H12
  await page.keyboard.type('지연된 입력'); await page.keyboard.press('Enter');
  await sleep(200);
  r = await agent('POST', `/api/docs/${opened.docId}/edit`, { edits: [{ op: 'insertRows', sheet: '매출', at: '1' }] });
  release(); await sleep(1500);
  await page.unroute('**/api/docs/*/ops');
  snap = await agent('GET', `/api/docs/${opened.docId}/snapshot`);
  const srv = snap.sheets[0].cells.find(([rr, cc]) => rr === 12 && cc === 7)?.[2]?.v;
  const brw = await page.evaluate(() => window.webExcel.doc.wb.sheet.get(12, 7)?.v);
  const stale = await page.evaluate(() => window.webExcel.doc.wb.sheet.get(11, 7)?.v);
  check('in-flight user edit transformed through agent insert', srv === '지연된 입력' && brw === '지연된 입력' && stale !== '지연된 입력', `server=${srv} browser=${brw} old=${stale}`);

  // 11) user inserts a row in the browser -> agent reads shifted formulas
  await page.evaluate(() => { const a = window.webExcel; a.selectRange({ r1: 1, c1: 0, r2: 1, c2: 16383 }); a.insert('row'); });
  await sleep(800);
  const rd = await agent('GET', `/api/docs/${opened.docId}/range?ref=${encodeURIComponent('매출!B12')}&what=formulas`);
  check('browser row insert reaches agent', rd.rows?.[0]?.[0] === '=SUM(B4:B11)', JSON.stringify(rd.rows));

  // 12) P7: agent moves the user's view and filters rows on screen
  r = await agent('POST', `/api/docs/${opened.docId}/goto`, { range: '비중!B3' });
  await page.waitForFunction(() => window.webExcel.doc.wb.sheet.name === '비중', null, { timeout: 4000 }).catch(() => {});
  check('agent goto moves preview to sheet/range', await page.evaluate(() => window.webExcel.doc.wb.sheet.name === '비중' && window.webExcel.view().active.r === 2 && window.webExcel.view().active.c === 1));
  await agent('POST', `/api/docs/${opened.docId}/view`, { zoom: 1.5 });
  await sleep(400);
  check('agent view sets zoom', await page.evaluate(() => window.webExcel.view().zoom) === 1.5);
  await page.evaluate(() => window.webExcel.activateSheet(0));
  r = await agent('POST', `/api/docs/${opened.docId}/edit`, { edits: [{ op: 'filterColumn', sheet: '매출', column: 'B', custom: [{ op: '>', val: '150' }] }] });
  await sleep(800);
  const hid = await page.evaluate(() => [...window.webExcel.doc.wb.sheet.filterHidden].length);
  check('agent filter hides rows in preview', hid > 0, `hidden=${hid} ${JSON.stringify(r.summary)}`);

  check('no page errors', errors.length === 0, errors.join(' | '));
  await agent('DELETE', `/api/docs/${opened.docId}`, { force: true });
  await browser.close();
  console.log(failures ? `\n${failures} FAILED` : '\nALL PASSED');
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
