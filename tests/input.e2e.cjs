// Typing after a click, frozen columns while scrolling, and pointing ranges in formulas.
// Starts its own server (port 8788).   node tests/input.e2e.cjs
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'monoonom', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-input-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, 'a.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

/** viewport point of the centre of cell (r, c), for real mouse clicks */
async function cellPoint(page, frame, r, c) {
  const pt = await frame.evaluate(([r, c]) => { const g = window.monoonom.grid; const h = g.host.getBoundingClientRect(); const rc = g.cellRect(r, c); return { x: h.left + rc.x + rc.w / 2, y: h.top + rc.y + rc.h / 2 }; }, [r, c]);
  if (frame === page.mainFrame()) return pt;
  const off = await page.evaluate(() => { const b = document.querySelector('iframe').getBoundingClientRect(); return { x: b.left, y: b.top }; });
  return { x: pt.x + off.x, y: pt.y + off.y };
}
const clickCell = async (page, frame, r, c, mods = []) => {
  const p = await cellPoint(page, frame, r, c);
  for (const k of mods) await page.keyboard.down(k);
  await page.mouse.click(p.x, p.y);
  for (const k of mods) await page.keyboard.up(k);
};
const cellOf = (frame, r, c) => frame.evaluate(([r, c]) => window.monoonom.doc.wb.sheet.get(r, c), [r, c]);
const edText = (frame) => frame.evaluate(() => window.monoonom.editor.ta.value);

(async () => {
  xl('serve', '--port', '8788'); xl('open', 'a.xlsx');
  const url = xl('url', '--bare').trim();
  const host = http.createServer((q, r) => { r.writeHead(200, { 'Content-Type': 'text/html' }); r.end(`<input id="chat" autofocus><br><iframe src="${url}" style="width:1100px;height:600px"></iframe>`); });
  await new Promise((r) => host.listen(8789, '127.0.0.1', r));
  const browser = await chromium.launch();
  const errors = [];
  try {
    // 1) click a cell, then type: goes into the cell (top-level tab and an iframe preview)
    for (const mode of ['top-level', 'iframe']) {
      const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
      page.on('pageerror', (e) => errors.push(e.message));
      let f;
      if (mode === 'top-level') { await page.goto(url); f = page.mainFrame(); }
      else { await page.goto('http://127.0.0.1:8789/'); await page.focus('#chat'); await sleep(300); f = page.frames().find((x) => x.url().startsWith('http://localhost')); }
      await f.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
      await clickCell(page, f, 12, 1);
      await page.keyboard.type(mode === 'iframe' ? '42' : '41'); await page.keyboard.press('Enter');
      await sleep(200);
      check(`${mode}: one click then typing enters the cell`, (await cellOf(f, 12, 1))?.v === (mode === 'iframe' ? 42 : 41), JSON.stringify(await cellOf(f, 12, 1)));
      check(`${mode}: focus is in the cell input after the click`, await f.evaluate(() => document.activeElement === window.monoonom.editor.ta));
      if (mode === 'iframe') check('iframe: keys did not go to the host page', (await page.evaluate(() => document.querySelector('#chat').value)) === '');
      await page.close();
    }

    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(url);
    const f = page.mainFrame();
    await f.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);

    // 2) frozen first column + horizontal scroll: no text of scrolled columns drawn in the frozen pane
    xl('write', 'C25', '셋째 열의 아주 긴 텍스트가 옆 칸으로 넘칩니다', '--doc', 'a.xlsx');
    await sleep(400);
    await page.evaluate(() => window.monoonom.exec('freezeCol'));
    const crossings = await page.evaluate(async () => {
      const g = window.monoonom.grid, m0 = g.metrics();
      const origRegion = g.drawRegion.bind(g), origText = g.drawText.bind(g);
      let region = null; const bad = [];
      g.drawRegion = (ctx, m, rg) => { region = rg; return origRegion(ctx, m, rg); };
      g.drawText = (ctx, m, d, r, c, ...rest) => { if (region && region.c1 < m.fc && c >= m.fc) bad.push(`${r},${c}`); if (region && region.c0 >= m.fc && c < m.fc) bad.push(`${r},${c}`); return origText(ctx, m, d, r, c, ...rest); };
      for (const x of [20, 60, 130, 260]) { g.scroller.scrollLeft = x; g.draw(); }
      g.drawRegion = origRegion; g.drawText = origText;
      return { fc: m0.fc, bad: [...new Set(bad)].slice(0, 8) };
    });
    check('frozen column: scrolled columns never drawn into the frozen pane', crossings.fc === 1 && crossings.bad.length === 0, JSON.stringify(crossings));
    await page.evaluate(() => { window.monoonom.exec('unfreeze'); window.monoonom.grid.scroller.scrollLeft = 0; });
    await sleep(200);

    // 3) pointing ranges in a formula
    await clickCell(page, f, 14, 0);                    // A15
    await page.keyboard.type('=SUM(');
    await clickCell(page, f, 1, 1);                     // B2
    await clickCell(page, f, 6, 1, ['Shift']);          // Shift+click B7
    check('Shift+click extends the pointed reference to a range', (await edText(f)) === '=SUM(B2:B7', await edText(f));
    await clickCell(page, f, 1, 3, ['Control']);        // Ctrl+click D2
    check('Ctrl+click adds another reference after a comma', (await edText(f)) === '=SUM(B2:B7,D2', await edText(f));
    await clickCell(page, f, 6, 3, ['Shift']);          // Shift+click D7
    check('…and Shift+click extends that one', (await edText(f)) === '=SUM(B2:B7,D2:D7', await edText(f));
    await page.keyboard.press('Enter');
    await sleep(200);
    check('formula is committed with the closing parenthesis', (await cellOf(f, 14, 0))?.f === '=SUM(B2:B7,D2:D7)', JSON.stringify(await cellOf(f, 14, 0)));
    // a reference typed by hand, then Shift+click
    await clickCell(page, f, 15, 0);                    // A16
    await page.keyboard.type('=AVERAGE(C2');
    await clickCell(page, f, 5, 2, ['Shift']);          // Shift+click C6
    check('Shift+click extends a typed reference', (await edText(f)) === '=AVERAGE(C2:C6', await edText(f));
    // Ctrl+drag adds a range
    await page.keyboard.type(')+SUM(');
    const a = await cellPoint(page, f, 1, 4), b = await cellPoint(page, f, 3, 4);
    await page.keyboard.down('Control'); await page.mouse.move(a.x, a.y); await page.mouse.down(); await page.mouse.move(b.x, b.y, { steps: 4 }); await page.mouse.up(); await page.keyboard.up('Control');
    check('Ctrl+click right after "(" inserts without a stray comma', (await edText(f)) === '=AVERAGE(C2:C6)+SUM(E2:E4', await edText(f));
    await page.keyboard.press('Escape');
    // plain click still replaces the reference just pointed (Excel behaviour)
    await clickCell(page, f, 16, 0);
    await page.keyboard.type('=');
    await clickCell(page, f, 1, 1); await clickCell(page, f, 2, 2);
    check('plain click replaces the reference just pointed', (await edText(f)) === '=C3', await edText(f));
    await page.keyboard.press('Escape');
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close(); host.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
