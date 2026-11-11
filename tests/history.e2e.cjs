// Change history panel: View > Change History lists every change (agent, you, another window) with
// where it happened; Undo takes back an agent's change in every window and on the server; a change
// whose cells were changed again asks first; a range link goes there; the filter shows agent changes
// only; the "Agent working" indicator opens the panel. Starts its own server (port 8885).
//   node tests/history.e2e.cjs
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'monoonom', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra && !ok ? ' — ' + extra : ''}`); if (!ok) fails++; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-hist-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

execFileSync('python3', ['-c', `
from openpyxl import Workbook
wb = Workbook(); s = wb.active; s.title = "Sales"
for r, (a, b) in enumerate([("Item", "Qty"), ("Apple", 3), ("Pear", 5)], 1): s.cell(r, 1, a); s.cell(r, 2, b)
wb.create_sheet("My Data")
wb.save(${JSON.stringify(path.join(ws, 'h.xlsx'))})
`]);

(async () => {
  xl('serve', '--port', '8885'); xl('open', 'h.xlsx');
  const url = xl('url', '--bare').trim();
  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [];
  try {
    const ctx = await browser.newContext({ viewport: { width: 1300, height: 800 } });
    const p = await ctx.newPage();
    p.on('pageerror', (e) => errors.push(e.message));
    p.on('console', (m) => { if (m.type() === 'error' && !/status of 409/.test(m.text())) errors.push(m.text()); });
    await p.goto(url);
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    const val = (r, c) => p.evaluate(([r, c]) => { const a = window.monoonom; return a.doc.wb.value(a.doc.wb.sheets[0], r, c); }, [r, c]);
    const items = () => p.evaluate(() => [...document.querySelectorAll('#history-panel .hp-item')].map((el) => ({
      who: el.querySelector('.hp-who').textContent, label: el.querySelector('.hp-label').textContent, where: el.querySelector('.hp-where')?.textContent || '',
      undo: !!el.querySelector('.hp-undo'), undone: el.classList.contains('undone'), batch: el.dataset.batch })));

    // the agent changes B2; the "Agent working" indicator opens the panel
    const out = xl('write', 'B2', '30');
    const agentBatch = /batch=(\w+)/.exec(out)[1];
    await p.waitForFunction(() => !document.getElementById('agent-busy').hidden, null, { timeout: 5000 });
    await p.click('#agent-busy'); await sleep(500);
    const gridW = await p.evaluate(() => document.getElementById('grid-host').offsetWidth);
    let it = await items();
    check('the "Agent working" indicator opens the panel; the grid makes room for it', await p.isVisible('#history-panel') && gridW <= 1300 - 320, `${gridW}`);
    check("…lists the agent's change (Korean: 에이전트, where, Undo button)", it.length === 1 && it[0].who === '에이전트' && /Sales!B2/.test(it[0].where) && it[0].undo && it[0].batch === agentBatch, JSON.stringify(it));
    check('…panel text is Korean', /변경 기록/.test(await p.textContent('#history-panel .hp-head')));

    // the user types in this window -> 나; the panel refreshes by itself
    await p.evaluate(() => { const a = window.monoonom; a.select(3, 0); a.focusGrid(); });
    await p.keyboard.type('Kiwi'); await p.keyboard.press('Enter'); await sleep(900);
    it = await items();
    check('your own edit shows up at the top as 나 (refreshed live)', it.length === 2 && it[0].who === '나' && /Sales!A4/.test(it[0].where), JSON.stringify(it));

    // a second window: the first window's edit is "another window" there
    const p2 = await ctx.newPage();
    await p2.goto(url); await p2.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    await p2.evaluate(() => window.monoonom.exec('history')); await sleep(600);
    const it2 = await p2.evaluate(() => [...document.querySelectorAll('#history-panel .hp-who')].map((x) => x.textContent));
    check('in another window the same edit is labelled 다른 창', it2[0] === '다른 창' && it2[1] === '에이전트', JSON.stringify(it2));

    // Undo the agent's change from the panel
    check('before undo: B2 is the agent\'s 30', (await val(1, 1)) === 30);
    await p.click(`#history-panel .hp-item[data-batch="${agentBatch}"] .hp-undo`); await sleep(900);
    it = await items();
    const read = JSON.parse(xl('read', 'B2', '--json')).rows[0][0];
    check("Undo restores B2 here, in the other window and on the server (3)", (await val(1, 1)) === 3 && read === 3
      && (await p2.evaluate(() => window.monoonom.doc.wb.value(window.monoonom.doc.wb.sheets[0], 1, 1))) === 3, `${await val(1, 1)} / ${read}`);
    const orig = it.find((x) => x.batch === agentBatch);
    check('…the change is marked undone (no Undo button) and the undo is listed as a new change by 나',
      orig.undone && !orig.undo && /되돌림/.test(orig.label) && it[0].who === '나' && /undo:/.test(it[0].label) && !it[0].undo, JSON.stringify(it));
    check('…status bar says so', /되돌림:/.test(await p.textContent('#status-msg')));

    // conflict: the agent writes C2, the user overwrites it, then Undo of the agent's change asks first
    const b2 = /batch=(\w+)/.exec(xl('write', 'C2', 'agent'))[1];
    await sleep(500);
    await p.evaluate(() => { const a = window.monoonom; a.select(1, 2); a.focusGrid(); });
    await p.keyboard.type('mine'); await p.keyboard.press('Enter'); await sleep(900);
    await p.click(`#history-panel .hp-item[data-batch="${b2}"] .hp-undo`);
    await p.waitForSelector('.modal-overlay .modal-body', { timeout: 5000 });
    const msg = await p.textContent('.modal-overlay .modal-body');
    check('undoing a change whose cells changed again asks first (Korean)', /이후/.test(msg) && /그래도 되돌리시겠습니까/.test(msg), msg);
    await p.click('.modal-overlay button[data-v="ok"], .modal-overlay .modal-buttons button:first-child').catch(async () => { await p.keyboard.press('Enter'); });
    await sleep(900);
    check('…confirming undoes it: C2 is empty again', (await val(1, 2)) == null || (await val(1, 2)) === '', String(await val(1, 2)));

    // a range link goes there; the filter shows agent changes only
    xl('write', "'My Data'!D5", 'x'); await sleep(900);
    await p.click('#history-panel .hp-item:first-child .hp-ref'); await sleep(300);
    const at = await p.evaluate(() => { const a = window.monoonom, v = a.view(); return { sheet: a.doc.wb.sheet.name, r: v.active.r, c: v.active.c }; });
    check('clicking a range in the panel goes there (sheet with a space in its name)', at.sheet === 'My Data' && at.r === 4 && at.c === 3, JSON.stringify(at));
    await p.selectOption('#history-panel .hp-filter', 'agent'); await sleep(200);
    it = await items();
    check('filter "에이전트": only agent changes', it.length === 3 && it.every((x) => x.who === '에이전트'), JSON.stringify(it.map((x) => x.who)));
    // View > Change History toggles it
    await p.evaluate(() => window.monoonom.exec('history')); await sleep(200);
    check('the View tab command closes it again; the grid gets its width back', !(await p.isVisible('#history-panel')) && (await p.evaluate(() => document.getElementById('grid-host').offsetWidth)) > gridW);
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch { /* ignore */ }
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
