// Font name / size pickers in the ribbon and the Format Cells dialog: a real dropdown (all items,
// whatever is in the box), fonts from this computer (/api/fonts), typing narrows the list, size
// checked like Excel (1-409, 0.5 steps). Starts its own server (port 8872).   node tests/fonts.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-fonts-'));
fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, 'sample.xlsx'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

(async () => {
  xl('serve', '--port', '8872'); xl('open', 'sample.xlsx');
  const sys = JSON.parse(xl('fonts', '--json'));
  check('xl.py fonts: the fonts on this computer', sys.fonts.length > 0, `${sys.fonts.length} fonts`);
  const some = sys.fonts.find((f) => /sans/i.test(f.name)) || sys.fonts[0];
  const browser = await chromium.launch();
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 800 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.webExcel?.doc?.sync); await sleep(300);
    await p.evaluate(() => { const a = window.webExcel; a.select(1, 1); a.focusGrid(); });
    const pop = () => p.evaluate(() => { const el = document.querySelector('.combo-pop'); return el ? { groups: [...el.querySelectorAll('.combo-group')].map((g) => g.textContent), items: [...el.querySelectorAll('.combo-item')].map((i) => i.textContent), current: el.querySelector('.combo-item.current')?.textContent || null } : null; });
    // font: the button opens everything, though the box says "Calibri"
    const before = await p.inputValue('#font-name');
    await p.click('#font-name + .combo-btn'); await sleep(400);
    let r = await pop();
    check('font list opens with all fonts, not only the ones matching the box', r && r.items.length >= sys.fonts.length + 1 && r.groups.includes('모든 글꼴') && r.groups[0] === '테마 글꼴', JSON.stringify(r && { n: r.items.length, groups: r.groups }));
    check('…the current font is marked', r && r.current === before, `${r && r.current} / ${before}`);
    check('…every system font is listed', r && sys.fonts.every((f) => r.items.includes(f.name)));
    // typing narrows it; Enter picks the highlighted one
    await p.fill('#font-name', some.name.slice(0, 4)); await sleep(200);
    r = await pop();
    check('typing narrows the list', r && r.items.length >= 1 && r.items.length < sys.fonts.length && r.items.every((x) => x.toLowerCase().includes(some.name.slice(0, 4).toLowerCase())), JSON.stringify(r && r.items.slice(0, 5)));
    await p.click(`.combo-pop .combo-item >> text="${some.name}"`); await sleep(300);
    const ff = await p.evaluate(() => window.webExcel.doc.wb.sheet.get(1, 1)?.s?.ff);
    check('clicking a font sets it on the selected cell', ff === some.name, String(ff));
    await p.click('#font-name + .combo-btn'); await sleep(300);
    r = await pop();
    check('…and it is listed under "이 문서에서 사용" and "최근 사용" next time', r && r.groups.includes('이 문서에서 사용') && r.groups.includes('최근 사용'), JSON.stringify(r && r.groups));
    await p.keyboard.press('Escape'); await sleep(100);
    check('Esc closes the list', !(await pop()));
    // size: all standard sizes, current marked, pick / type / reject
    await p.click('#font-size + .combo-btn'); await sleep(200);
    r = await pop();
    check('size list: all 16 standard sizes, current 11 marked', r && r.items.length === 16 && r.items[0] === '8' && r.items[15] === '72' && r.current === '11', JSON.stringify(r));
    await p.click('.combo-pop .combo-item >> text="14"'); await sleep(200);
    check('picking 14 sets the size', (await p.evaluate(() => window.webExcel.doc.wb.sheet.get(1, 1)?.s?.fs)) === 14);
    await p.fill('#font-size', '10.3'); await p.keyboard.press('Enter'); await sleep(200);
    check('typing 10.3 gives 10.5 (0.5 steps)', (await p.evaluate(() => window.webExcel.doc.wb.sheet.get(1, 1)?.s?.fs)) === 10.5);
    await p.fill('#font-size', '500'); await p.keyboard.press('Enter'); await sleep(200);
    const toast = await p.evaluate(() => [...document.querySelectorAll('.toast')].map((t) => t.textContent).join(' | '));
    check('typing 500 is refused with one message; the size stays', (await p.evaluate(() => window.webExcel.doc.wb.sheet.get(1, 1)?.s?.fs)) === 10.5 && (toast.match(/1에서 409/g) || []).length === 1, toast);
    // keyboard: ArrowDown opens, arrows move, Enter picks
    await p.click('#font-size'); await p.keyboard.press('ArrowDown'); await sleep(150);
    await p.keyboard.press('ArrowDown'); await p.keyboard.press('Enter'); await sleep(200);
    const kb = await p.evaluate(() => window.webExcel.doc.wb.sheet.get(1, 1)?.s?.fs);
    check('keyboard: ArrowDown opens, ArrowDown moves past the current size, Enter picks it', kb === 11, String(kb));
    // the agent sees it
    check('agent reads the style', /"ff": "[^"]+"/.test(xl('read', 'B2', '--what', 'all', '--json')));
    // Format Cells dialog: the font field is the same picker
    await p.evaluate(() => window.webExcel.exec('formatCells'));
    await p.waitForSelector('.fmt-dialog [data-tab="font"]');
    await p.click('.fmt-dialog [data-tab="font"]'); await sleep(150);
    await p.click('.fmt-dialog .combo-btn'); await sleep(300);
    r = await pop();
    check('Format Cells: the font field opens the same list', r && r.groups.includes('모든 글꼴') && r.items.length >= sys.fonts.length, JSON.stringify(r && r.groups));
    if (process.env.SHOT) await p.screenshot({ path: process.env.SHOT });
    await p.keyboard.press('Escape');
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch {}
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch {} process.exit(1); });
