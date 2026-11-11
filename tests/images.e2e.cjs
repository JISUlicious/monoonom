// Pictures: a file's picture is shown in the preview; pasting a picture (clipboard with only an image)
// adds one at the active cell; text with a picture (Excel's cell copy) still pastes as cells; move,
// resize, Delete and undo; the agent sees them (image ls); saving writes real pictures; row inserts
// move them. Starts its own server (port 8880).   node tests/images.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-img-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const py = (code) => execFileSync('python3', ['-c', code], { encoding: 'utf8' });

py(`
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as P
b = io.BytesIO(); P.new("RGB", (60, 40), (200, 40, 40)).save(b, "PNG"); b.seek(0)
wb = Workbook(); sh = wb.active; sh.title = "S"; sh["A1"] = "title"
img = Image(b); sh.add_image(img, "C3")
wb.save(${JSON.stringify(path.join(ws, 'pic.xlsx'))})
`);
const pics = (file) => JSON.parse(py(`
import json, openpyxl
ws = openpyxl.load_workbook(${JSON.stringify(path.join(ws, '__F__'))}.replace('__F__', ${JSON.stringify(file)}))['S']
E = 9525
print(json.dumps([[i.anchor._from.col, i.anchor._from.row, round(i.anchor.ext.width / E), round(i.anchor.ext.height / E), i.format] for i in ws._images]))`));

(async () => {
  xl('serve', '--port', '8880'); xl('open', 'pic.xlsx');
  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [];
  try {
    const p = await browser.newPage({ viewport: { width: 1300, height: 800 } });
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(xl('url', '--bare').trim());
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    await p.waitForFunction(() => document.querySelector('.image-box img')?.naturalWidth > 0, null, { timeout: 5000 }).catch(() => {});
    const shown = await p.evaluate(() => { const el = document.querySelector('.image-box'), im = el?.querySelector('img'); return el && { w: el.offsetWidth, h: el.offsetHeight, nw: im.naturalWidth }; });
    check('the file\'s picture is shown (60 x 40)', shown && shown.w === 60 && shown.h === 40 && shown.nw === 60, JSON.stringify(shown));
    check('…no "not shown" warning any more', !(await p.evaluate(() => (window.monoonom.doc.wb.warnings || []).join(' '))).includes('mage'));

    // paste a picture only (a screenshot) at E10
    const paste = (withText) => p.evaluate(async (withText) => {
      const a = window.monoonom; a.select(9, 4); a.focusGrid();
      const c = document.createElement('canvas'); c.width = 120; c.height = 80;
      const g = c.getContext('2d'); g.fillStyle = '#2a7'; g.fillRect(0, 0, 120, 80);
      const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
      const dt = new DataTransfer();
      dt.items.add(new File([blob], 'shot.png', { type: 'image/png' }));
      if (withText) dt.setData('text/plain', 'a\tb');
      document.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true }));
    }, withText);
    await paste(false);
    await p.waitForFunction(() => window.monoonom.doc.wb.sheet.images.length === 2, null, { timeout: 5000 }).catch(() => {});
    let st = await p.evaluate(() => { const a = window.monoonom, ims = a.doc.wb.sheet.images; return { n: ims.length, last: ims.at(-1)?.anchor, sel: a.images.selected === ims.at(-1)?.id, status: document.getElementById('status-msg').textContent }; });
    check('pasting a picture adds it at the active cell (E10), 120 x 80, selected', st.n === 2 && st.last?.c === 4 && st.last?.r === 9 && st.last.w === 120 && st.last.h === 80 && st.sel, JSON.stringify(st));
    check('…status bar says so (Korean)', /그림을 붙여넣었습니다/.test(st.status), st.status);
    await sleep(500);
    check('…the agent sees it (image ls)', /S\s+im\w+\s+png\s+@E10\s+120x80 px/.test(xl('image', 'ls')), xl('image', 'ls'));
    // text and a picture together (Excel's copy): cells
    await paste(true); await sleep(400);
    st = await p.evaluate(() => { const a = window.monoonom, sh = a.doc.wb.sheet; return { n: sh.images.length, e10: sh.get(9, 4)?.v, f10: sh.get(9, 5)?.v }; });
    check('text with a picture pastes as cells, not as a picture', st.n === 2 && st.e10 === 'a' && st.f10 === 'b', JSON.stringify(st));
    // move the pasted picture by dragging, resize from a corner
    const box = async () => p.evaluate(() => { const a = window.monoonom, id = a.doc.wb.sheet.images.at(-1).id; const el = a.images.instances.get(id).el.getBoundingClientRect(); return { x: el.x, y: el.y, w: el.width, h: el.height }; });
    let b = await box();
    await p.mouse.move(b.x + 30, b.y + 30); await p.mouse.down(); await p.mouse.move(b.x + 30 + 128, b.y + 30 + 40, { steps: 5 }); await p.mouse.up(); await sleep(300);
    let a2 = await p.evaluate(() => window.monoonom.doc.wb.sheet.images.at(-1).anchor);
    check('dragging moves it (two columns right, two rows down)', a2.c === 6 && a2.r === 11, JSON.stringify(a2));
    b = await box();
    await p.evaluate(() => { const a = window.monoonom; a.images.select(a.doc.wb.sheet.images.at(-1).id); }); await sleep(100);
    await p.mouse.move(b.x + b.w - 1, b.y + b.h - 1); await p.mouse.down(); await p.mouse.move(b.x + b.w + 59, b.y + b.h + 9, { steps: 4 }); await p.mouse.up(); await sleep(300);
    a2 = await p.evaluate(() => window.monoonom.doc.wb.sheet.images.at(-1).anchor);
    check('dragging the corner resizes it (180 x 90)', a2.w === 180 && a2.h === 90, JSON.stringify(a2));
    // Delete, then undo
    await p.keyboard.press('Delete'); await sleep(300);
    const afterDel = await p.evaluate(() => window.monoonom.doc.wb.sheet.images.length);
    await p.keyboard.press('Control+z'); await sleep(400);
    const afterUndo = await p.evaluate(() => window.monoonom.doc.wb.sheet.images.length);
    check('Delete removes the selected picture; Ctrl+Z brings it back', afterDel === 1 && afterUndo === 2, `${afterDel} / ${afterUndo}`);
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  await sleep(300);
  xl('save', '--as', 'out.xlsx');
  let r = pics('out.xlsx');
  check('saved: the file\'s picture where it was (C3, 60x40) and the pasted one (G12, 180x90) as real pictures',
    JSON.stringify(r) === JSON.stringify([[2, 2, 60, 40, 'png'], [6, 11, 180, 90, 'png']]), JSON.stringify(r));
  xl('rows', 'insert', '1', '--count', '2');
  xl('save', '--as', 'out2.xlsx');
  r = pics('out2.xlsx');
  check('inserting rows above moves both pictures down', r.map((x) => x[1]).join() === '4,13', JSON.stringify(r));
  try { xl('stop', '--force'); } catch { /* ignore */ }
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
