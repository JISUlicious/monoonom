// Argument hints while typing a formula, and following the agent's work (on/off).
// Starts its own server (port 8792).   node tests/follow.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-follow-'));
for (const n of ['a', 'b']) fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, `${n}.xlsx`));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

(async () => {
  xl('serve', '--port', '8792'); xl('open', 'a.xlsx');
  const url = xl('url', '--bare').trim();
  const browser = await chromium.launch();
  const errors = [];
  try {
    const ctx = await browser.newContext();
    const p = await ctx.newPage({ viewport: { width: 1200, height: 700 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(url);
    await p.waitForFunction(() => window.webExcel?.doc?.sync); await sleep(300);
    const st = () => p.evaluate(() => { const a = window.webExcel; return { doc: a.doc?.wb.name, sheet: a.doc?.wb.sheet.name, vis: a.grid.lastVisible }; });

    // ---- 1) argument hints
    const hint = async (text) => {
      await p.evaluate(() => { const a = window.webExcel; a.select(12, 0); a.focusGrid(); });
      await p.keyboard.type(text); await sleep(100);
      const r = await p.evaluate(() => { const h = document.querySelector('.fn-hint'); return h && !h.classList.contains('hidden') ? { text: h.textContent, bold: [...h.querySelectorAll('b')].map((b) => b.textContent) } : null; });
      await p.keyboard.press('Escape');
      return r;
    };
    let h = await hint('=SUM(1,2,');
    check('hint: third SUM argument is named [number3]', h && JSON.stringify(h.bold) === '["[number3]"]', JSON.stringify(h));
    h = await hint('=SUMIFS(A1:A3,B1:B3,"x",');
    check('hint: SUMIFS 4th argument is [criteria_range2]', h && h.bold[0] === '[criteria_range2]', JSON.stringify(h));
    h = await hint('=NORM.DIST(1,');
    check('hint: functions that had no hint before (NORM.DIST) now show one', h && h.bold[0] === 'mean' && /정규 분포/.test(h.text), JSON.stringify(h));
    h = await hint('=VLOOKUP(A1,{1,2,3},');
    check('hint: commas inside {array} constants are not counted', h && h.bold[0] === 'col_index_num', JSON.stringify(h));
    h = await hint('=IF(A1>0,SUM(1,');
    check('hint: innermost function wins', h && h.text.startsWith('SUM(') && h.bold[0] === '[number2]', JSON.stringify(h));

    // ---- 2) follow on (default)
    check('follow is on by default', await p.evaluate(() => window.webExcel.follow && document.getElementById('follow-agent').checked));
    xl('write', '비중!A3', '에이전트', '--doc', 'a.xlsx');
    await p.waitForFunction(() => window.webExcel.doc.wb.sheet.name === '비중', null, { timeout: 4000 }).catch(() => {});
    check('on: agent works on another sheet -> the view switches to it', (await st()).sheet === '비중', JSON.stringify(await st()));
    check('…the agent-busy indicator shows', await p.evaluate(() => !document.getElementById('agent-busy').hidden));
    xl('write', '비중!B300', '멀리', '--doc', 'a.xlsx');
    await sleep(800);
    let s = await st();
    check('on: a far-away cell is scrolled into view', s.vis.r0 <= 299 && 299 <= s.vis.r1, JSON.stringify(s.vis));
    check('…and outlined as the agent\'s cell', await p.evaluate(() => (window.webExcel.doc.agentMarks || []).some((m) => m.r1 === 299 && m.c1 === 1)));
    xl('sheet', 'add', '새시트', '--doc', 'a.xlsx');
    await p.waitForFunction(() => window.webExcel.doc.wb.sheet.name === '새시트', null, { timeout: 4000 }).catch(() => {});
    check('on: a sheet the agent adds is shown', (await st()).sheet === '새시트', JSON.stringify(await st()));
    xl('open', 'b.xlsx');
    await p.waitForFunction(() => window.webExcel.doc.wb.name === 'b.xlsx', null, { timeout: 4000 }).catch(() => {});
    await p.evaluate(() => window.webExcel.activateDoc(0));
    xl('write', 'C5', 'b에서', '--doc', 'b.xlsx');
    await p.waitForFunction(() => window.webExcel.doc.wb.name === 'b.xlsx', null, { timeout: 4000 }).catch(() => {});
    check('on: agent edits another open workbook -> it is shown', (await st()).doc === 'b.xlsx', JSON.stringify(await st()));
    // typing protects the user
    await p.evaluate(() => { const a = window.webExcel; a.activateSheet(0); a.select(20, 0); a.focusGrid(); });
    await p.keyboard.type('입력 중');
    xl('write', '비중!A2', 'x', '--doc', 'b.xlsx');
    await sleep(800);
    s = await st();
    check('on, but the user is typing: the view stays and the sheet tab is marked', s.sheet === '매출' && await p.evaluate(() => window.webExcel.editor.active && !!document.querySelector('.sheet-tab .st-new')), JSON.stringify(s));
    await p.keyboard.press('Escape');

    // ---- 3) follow off
    await p.click('#follow-agent');
    check('toggle turns follow off', await p.evaluate(() => !window.webExcel.follow));
    await p.evaluate(() => window.webExcel.activateSheet(0));
    xl('write', '비중!D2', '끔', '--doc', 'b.xlsx');
    await sleep(800);
    s = await st();
    check('off: the view stays on the user\'s sheet', s.sheet === '매출' && s.doc === 'b.xlsx', JSON.stringify(s));
    check('…the changed sheet\'s tab is marked', await p.evaluate(() => [...document.querySelectorAll('.sheet-tab')].some((t) => t.textContent.startsWith('비중') && t.querySelector('.st-new'))));
    xl('write', '매출!E3', '여기', '--doc', 'b.xlsx');
    await sleep(500);
    check('off: agent cells on the visible sheet are outlined', await p.evaluate(() => (window.webExcel.doc.agentMarks || []).some((m) => m.sheet.name === '매출' && m.r1 === 2 && m.c1 === 4)));
    await p.evaluate(() => window.webExcel.activateSheet(window.webExcel.doc.wb.sheets.findIndex((x) => x.name === '비중')));
    check('…opening the marked sheet clears its marker', await p.evaluate(() => !document.querySelector('.sheet-tab.on .st-new')));
    if (process.env.SHOT) { await p.evaluate(() => window.webExcel.activateSheet(0)); xl('write', '매출!B2', '7', '--doc', 'b.xlsx'); await sleep(400); await p.screenshot({ path: process.env.SHOT }); }

    // ---- 4) the setting survives a reload; the agent can switch it
    await p.reload(); await p.waitForFunction(() => window.webExcel?.started);
    check('off survives a reload', await p.evaluate(() => !window.webExcel.follow && !document.getElementById('follow-agent').checked));
    xl('view', '--follow', 'on', '--doc', 'b.xlsx');
    check('xl.py view --follow on turns it on', await p.waitForFunction(() => window.webExcel.follow && document.getElementById('follow-agent').checked, null, { timeout: 4000 }).then(() => true).catch(() => false));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
