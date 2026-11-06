// Which document a preview shows: agent opens/reopens, reloads, new windows, background changes,
// and no automatic blank workbook. Starts its own server (port 8783).   node tests/focus.e2e.cjs
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

const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-focus-'));
for (const n of ['a', 'b']) fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, `${n}.xlsx`));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const docs = () => JSON.parse(xl('docs', '--json')).map((d) => d.name).sort();
const state = (p) => p.evaluate(() => ({
  tabs: window.monoonom.docs.map((d) => d.wb.name),
  shown: window.monoonom.doc?.wb.name || null,
  marked: window.monoonom.docs.filter((d) => d.unseen).map((d) => d.wb.name),
  empty: !document.getElementById('empty-state')?.hidden && !window.monoonom.docs.length,
}));
const until = (p, fn, arg) => p.waitForFunction(fn, arg, { timeout: 5000 }).then(() => true).catch(() => false);
const shownIs = (p, name) => until(p, (n) => window.monoonom.doc?.wb.name === n, name);

(async () => {
  xl('serve', '--port', '8783');
  const url = xl('url', '--bare').trim();
  const browser = await browser_();
  async function browser_() { return require('./korean.cjs')(await chromium.launch()); }
  const errors = [];
  try {
    const ctx = await browser.newContext();
    const p = await ctx.newPage();
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(url);
    await p.waitForFunction(() => window.monoonom?.started);
    let s = await state(p);
    check('empty server: start screen, no workbook created', s.empty && docs().length === 0, JSON.stringify(s));

    xl('open', 'a.xlsx');
    check('agent opens a: shown', await shownIs(p, 'a.xlsx'));
    s = await state(p);
    check('…no blank workbook next to it', JSON.stringify(s.tabs) === '["a.xlsx"]' && JSON.stringify(docs()) === '["a.xlsx"]', JSON.stringify(s.tabs));
    xl('open', 'b.xlsx');
    check('agent opens b: shown', await shownIs(p, 'b.xlsx'));
    await p.evaluate(() => window.monoonom.activateDoc(window.monoonom.docs.findIndex((d) => d.wb.name === 'a.xlsx')));
    xl('open', 'b.xlsx');
    check('agent opens b again while the user is on a: b is shown', await shownIs(p, 'b.xlsx'));

    // the user is typing in a cell: the agent does not pull the screen away, the tab is marked
    await p.evaluate(() => window.monoonom.activateDoc(window.monoonom.docs.findIndex((d) => d.wb.name === 'a.xlsx')));
    await p.evaluate(() => { const a = window.monoonom; a.select(30, 0); a.focusGrid(); });
    await p.keyboard.type('입력 중');
    xl('open', 'b.xlsx');
    await sleep(800);
    s = await state(p);
    check('typing: stays on a, keeps the half-typed text, marks b', s.shown === 'a.xlsx' && s.marked.includes('b.xlsx') && await p.evaluate(() => window.monoonom.editor.active && window.monoonom.editor.ta.value === '입력 중'), JSON.stringify(s));
    await p.keyboard.press('Escape');

    // agent edits b while a is on screen, with follow-the-agent off (on would switch to b, see
    // follow.e2e.cjs): b's tab gets a marker; clicking it clears the marker
    await p.evaluate(() => { window.monoonom.setFollow(false); window.monoonom.docs.forEach((d) => { d.unseen = false; }); });
    xl('write', 'A40', '에이전트', '--doc', 'b.xlsx');
    check('background change: b marked, a still shown', await until(p, () => window.monoonom.docs.find((d) => d.wb.name === 'b.xlsx')?.unseen) && (await state(p)).shown === 'a.xlsx');
    check('…marker is drawn on the tab', await p.evaluate(() => !!document.querySelector('.doc-tab .dt-new')));
    await p.evaluate(() => window.monoonom.activateDoc(window.monoonom.docs.findIndex((d) => d.wb.name === 'b.xlsx')));
    check('…and cleared when the tab is opened', !(await state(p)).marked.length);
    await p.evaluate(() => window.monoonom.setFollow(true));

    // reload: this window keeps the document it showed
    await p.evaluate(() => window.monoonom.activateDoc(window.monoonom.docs.findIndex((d) => d.wb.name === 'a.xlsx')));
    await p.reload(); await p.waitForFunction(() => window.monoonom?.started);
    s = await state(p);
    check('reload keeps the document this window showed (a)', s.shown === 'a.xlsx' && s.tabs.length === 2, JSON.stringify(s));

    // a fresh window shows what was used last (the agent's goto on b)
    xl('goto', 'B2', '--doc', 'b.xlsx');
    check('goto moves the open window to b', await shownIs(p, 'b.xlsx'));
    await sleep(800); // let window 1's selection reports (presence also counts as "in use") arrive
    const ctx2 = await browser.newContext();
    const p2 = await ctx2.newPage();
    p2.on('pageerror', (e) => errors.push('p2: ' + e.message));
    await p2.goto(url); await p2.waitForFunction(() => window.monoonom?.started);
    check('new window shows the most recently used document (b)', (await state(p2)).shown === 'b.xlsx', (await state(p2)).shown);

    // a user creating a workbook in window 2 does not pull window 1 to it
    await p.evaluate(() => window.monoonom.activateDoc(window.monoonom.docs.findIndex((d) => d.wb.name === 'a.xlsx')));
    await p2.evaluate(() => window.monoonom.newDoc());
    await until(p, () => window.monoonom.docs.length === 3);
    s = await state(p);
    check('other window\'s new workbook: background tab here, a still shown', s.tabs.length === 3 && s.shown === 'a.xlsx', JSON.stringify(s));
    await ctx2.close();

    // closing every tab: start screen, nothing created on the server
    for (const d of JSON.parse(xl('docs', '--json'))) xl('close', '--doc', d.docId, '--force');
    check('all closed: start screen in the window', await until(p, () => !window.monoonom.docs.length && !document.getElementById('empty-state').hidden));
    await sleep(500);
    check('…and no blank workbook was created on the server', docs().length === 0, JSON.stringify(docs()));
    await p.keyboard.press('Control+b'); await p.keyboard.press('Delete'); await p.keyboard.press('Control+Tab');
    check('keys on the start screen do nothing harmful', errors.length === 0, errors.join('; '));
    xl('open', 'a.xlsx');
    check('agent opens a file: start screen goes away', await shownIs(p, 'a.xlsx') && await p.evaluate(() => document.getElementById('empty-state').hidden));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
