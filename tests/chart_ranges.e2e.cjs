// Selecting a chart outlines its data ranges on the grid (name green, X purple, Y blue) and lists
// them per series in a panel; clicking a series row shows only that series. The agent sees the same
// ranges in `chart ls`.   Starts its own server (port 8840).   node tests/chart_ranges.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-cranges-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, 'sample.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

(async () => {
  xl('serve', '--port', '8840'); xl('open', 'sample.xlsx');
  // agent side: per-series name / X / Y
  const ls = xl('chart', 'ls');
  check('chart ls: one line per series with name, X and Y ranges',
    /계열 1: 이름 '매출'!B1  X '매출'!\$A\$2:\$A\$7  Y '매출'!\$B\$2:\$B\$7/.test(ls) && /계열 4 \(line\)/.test(ls), ls);
  const browser = await chromium.launch();
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 760 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.wb.sheet.charts.length > 0); await sleep(300);
    const st = () => p.evaluate(() => {
      const a = window.monoonom, pnl = document.querySelector('.chart-ranges');
      return {
        ranges: a.grid.chartRanges.map((h) => ({ c: h.color, r: [h.range.r1, h.range.c1, h.range.r2, h.range.c2], dim: !!h.dim })),
        panel: pnl && pnl.style.display !== 'none' ? pnl.textContent : null,
        rows: pnl ? pnl.querySelectorAll('.cr-row').length : 0,
      };
    });
    check('nothing outlined before a chart is selected', (await st()).ranges.length === 0);
    await p.evaluate(() => { const a = window.monoonom; a.charts.select(a.doc.wb.sheet.charts[0].id); });
    await sleep(200);
    let s = await st();
    const has = (c, r) => s.ranges.some((h) => h.c === c && h.r.join() === r.join());
    check('series names outlined green (B1..E1)', [1, 2, 3, 4].every((c) => has('#008000', [0, c, 0, c])), JSON.stringify(s.ranges));
    check('categories outlined purple (A2:A7), once though 3 series share them', s.ranges.filter((h) => h.c === '#7030A0').length === 1 && has('#7030A0', [1, 0, 6, 0]));
    check('values outlined blue (B2:B7 … E2:E7)', [1, 2, 3, 4].every((c) => has('#0070C0', [1, c, 6, c])));
    check('panel lists 4 series with their refs', s.rows === 4 && s.panel.includes("Y(값) '매출'!$C$2:$C$7") && s.panel.includes('(line)') && s.panel.includes('제품A') && s.panel.includes('합계'), s.panel);
    const frames = await p.evaluate(() => new Promise((res) => {
      const g = window.monoonom.grid, d = g.draw.bind(g); let n = 0;
      g.draw = () => { n++; d(); }; setTimeout(() => { g.draw = d; res(n); }, 500);
    }));
    check('a selected chart does not redraw the grid every frame', frames <= 1, `${frames} draws in 500 ms`);
    if (process.env.SHOT) await p.screenshot({ path: process.env.SHOT });
    // focus one series
    await p.click('.chart-ranges .cr-row[data-series="1"]');
    await sleep(150);
    s = await st();
    const live = s.ranges.filter((h) => !h.dim).map((h) => h.r[1]);
    check('clicking a series row: only that series stays bright (name, X, Y)', live.length === 3 && live.filter((c) => c === 2).length === 2 && live.includes(0) && s.ranges.some((h) => h.dim), JSON.stringify(s.ranges));
    await p.click('.chart-ranges .cr-row[data-series="1"]');
    await sleep(150);
    check('clicking it again: all series bright', (await st()).ranges.every((h) => !h.dim));
    // deselect
    await p.evaluate(() => window.monoonom.charts.deselect()); await sleep(150);
    s = await st();
    check('deselect: outlines and panel go away', s.ranges.length === 0 && s.panel === null, JSON.stringify(s));
    // pie on another sheet
    await p.evaluate(() => { const a = window.monoonom; a.activateSheet(a.doc.wb.sheets.findIndex((x) => x.name === '비중')); });
    await sleep(200);
    await p.evaluate(() => { const a = window.monoonom; a.charts.select(a.doc.wb.sheet.charts[0].id); }); await sleep(200);
    s = await st();
    check('pie: name, categories and values outlined', s.ranges.length === 3 && s.rows === 1, JSON.stringify(s));

    // ---- switch rows / columns
    await p.evaluate(() => window.monoonom.activateSheet(0)); await sleep(150);
    await p.evaluate(() => { const a = window.monoonom; a.charts.select(a.doc.wb.sheet.charts[0].id); }); await sleep(150);
    await p.click('.chart-ranges [data-switch]'); await sleep(400);
    const sw = await p.evaluate(() => window.monoonom.doc.wb.sheet.charts[0].series.map((x) => [x.nameRef, x.cat, x.val].map((t) => t && t.replace(/'/g, '')).join(' ')));
    check('UI 행/열 전환: 6 series (one per month), names A2..A7, X = B1:E1, Y = a row',
      sw.length === 6 && sw[0] === '매출!$A$2 매출!$B$1:$E$1 매출!$B$2:$E$2' && sw[5] === '매출!$A$7 매출!$B$1:$E$1 매출!$B$7:$E$7', JSON.stringify(sw));
    s = await st();
    check('…the outlines follow (names A2:A7 green, X B1:E1 purple)', has('#008000', [1, 0, 1, 0]) && has('#7030A0', [0, 1, 0, 4]) && has('#0070C0', [6, 1, 6, 4]), JSON.stringify(s.ranges));
    const agentSees = xl('chart', 'ls');
    check('…the server (agent) sees the switched chart', /계열 6: 이름 '?매출'?!\$A\$7 {2}X '?매출'?!\$B\$1:\$E\$1 {2}Y '?매출'?!\$B\$7:\$E\$7/.test(agentSees), agentSees);
    const cid = await p.evaluate(() => window.monoonom.doc.wb.sheet.charts[0].id);
    xl('chart', 'set', cid, '--switch');
    await p.waitForFunction(() => window.monoonom.doc.wb.sheet.charts[0].series.length === 4, null, { timeout: 4000 }).catch(() => {});
    const back = await p.evaluate(() => window.monoonom.doc.wb.sheet.charts[0].series.map((x) => [x.nameRef, x.cat, x.val].map((t) => t && t.replace(/'/g, '')).join(' ')));
    check('agent `chart set --switch` switches it back; the preview shows the original ranges',
      back.length === 4 && back[0] === '매출!$B$1 매출!$A$2:$A$7 매출!$B$2:$B$7' && back[3] === '매출!$E$1 매출!$A$2:$A$7 매출!$E$2:$E$7', JSON.stringify(back));
    // the same rules in JS (UI) and Python (agent)
    const cases = [
      [{ val: 'S!$B$2:$B$5', cat: 'S!$A$2:$A$5', nameRef: 'S!$B$1' }, { val: 'S!$C$2:$C$5', cat: 'S!$A$2:$A$5', nameRef: 'S!$C$1' }],
      [{ val: 'S!$B$2:$D$2', cat: 'S!$B$1:$D$1', nameRef: 'S!$A$2' }, { val: 'S!$B$3:$D$3', nameRef: 'S!$A$3' }],
      [{ val: 'S!$B$2:$B$4', name: '리터럴' }],
      [{ val: 'S!$B$2:$B$4' }, { val: 'S!$D$2:$D$4' }],
      [{ val: 'S!$B$2:$B$4', cat: 'S!$A$2:$A$4' }, { val: 'S!$C$2:$C$4', cat: 'S!$E$2:$E$4' }],
      [{ val: 'S!$B$2:$B$4', cat: 'S!$A$2:$A$3' }],
    ];
    const js = await p.evaluate((cs) => cs.map((c) => { try { return window.monoonom.charts.constructor.switchRowCol(c); } catch (e) { return 'ERR ' + e.message; } }), cases);
    const py = JSON.parse(execFileSync('python3', ['-c', `
import json, sys; sys.path.insert(0, ${JSON.stringify(path.join(ROOT, '.agents', 'skills', 'monoonom', 'scripts'))})
from server.agent_more import switch_row_col
out = []
for c in json.loads(sys.argv[1]):
    try: out.append(switch_row_col(c))
    except Exception as e: out.append('ERR ' + str(getattr(e, 'message', None) or e.args[-1]))
print(json.dumps(out, ensure_ascii=False))`, JSON.stringify(cases)], { encoding: 'utf8' }));
    const norm = (r) => JSON.stringify(typeof r === 'string' ? r.startsWith('ERR') : r.map((x) => [x.nameRef, x.cat, x.val]));
    const same = js.every((r, i) => norm(r) === norm(py[i]));
    check('JS and Python switch the same way (6 cases incl. errors)', same && typeof js[3] === 'string' && typeof js[4] === 'string' && typeof js[5] === 'string', JSON.stringify({ js, py }));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
