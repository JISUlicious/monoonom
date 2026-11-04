// End-to-end test (requires: npm i playwright, and the server running on :8765)
//   python run.py --port 8765 --root . &   node tests/e2e.cjs
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = process.env.BASE || 'http://127.0.0.1:8765/';
const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

let failures = 0;
const check = (name, cond, extra = '') => { console.log(`${cond ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!cond) failures++; };

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 860 } });
  await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: BASE });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  // a 409 is how the server asks before replacing a file (the overwrite prompt is tested below)
  page.on('console', (m) => { if (m.type() === 'error' && !/status of 409/.test(m.text())) errors.push(m.text()); });
  await page.goto(BASE);
  await page.waitForFunction(() => window.webExcel?.doc);
  const val = (r, c, sheet) => page.evaluate(([r, c, s]) => { const a = window.webExcel; const sh = s == null ? a.doc.wb.sheet : a.doc.wb.sheets[s]; const v = a.doc.wb.value(sh, r, c); return v && v.err ? v.err : v; }, [r, c, sheet]);
  const cell = (r, c) => page.evaluate(([r, c]) => window.webExcel.doc.wb.sheet.get(r, c) || null, [r, c]);
  const sel = (r, c) => page.evaluate(([r, c]) => window.webExcel.select(r, c), [r, c]);
  const box = async (r, c) => page.evaluate(([r, c]) => { const a = window.webExcel; const rc = a.grid.cellRect(r, c); const h = a.gridHost.getBoundingClientRect(); return { x: h.left + rc.x + rc.w / 2, y: h.top + rc.y + rc.h / 2, w: rc.w, h: rc.h, l: h.left + rc.x, t: h.top + rc.y }; }, [r, c]);

  // ---- typing & formulas
  await page.keyboard.type('10'); await page.keyboard.press('Enter');
  await page.keyboard.type('20'); await page.keyboard.press('Enter');
  await page.keyboard.type('=sum(a1:a2'); await page.keyboard.press('Enter');
  check('formula with auto-closed paren', (await val(2, 0)) === 30);
  check('formula normalized to upper case', (await cell(2, 0)).f === '=SUM(A1:A2)');
  // IME-like insertText (Hangul)
  await sel(0, 1);
  await page.keyboard.insertText('한글 입력');
  await page.keyboard.press('Enter');
  check('hangul text via input event', (await val(0, 1)) === '한글 입력');

  // ---- point mode: type =, click A1, type +, arrow up
  await sel(4, 0);
  await page.keyboard.type('=');
  const a1 = await box(0, 0);
  await page.mouse.click(a1.x, a1.y);
  await page.keyboard.type('*2+');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Enter');
  check('point mode inserts references', (await cell(4, 0)).f === '=A1*2+A4', (await cell(4, 0)).f);

  // ---- F4 absolute toggle
  await sel(5, 0);
  await page.keyboard.type('=A1'); await page.keyboard.press('F4'); await page.keyboard.press('Enter');
  check('F4 toggles absolute ref', (await cell(5, 0)).f === '=$A$1');

  // ---- styling shortcuts + undo/redo
  await sel(0, 0);
  await page.keyboard.press('Control+b');
  check('Ctrl+B bold', (await cell(0, 0)).s?.b === true);
  await page.keyboard.press('Control+z');
  check('Ctrl+Z undo', !(await cell(0, 0)).s?.b);
  await page.keyboard.press('Control+y');
  check('Ctrl+Y redo', (await cell(0, 0)).s?.b === true);

  // ---- copy / paste with relative formula translation
  await sel(2, 0);
  await page.keyboard.press('Control+c');
  await sel(2, 2);
  await page.keyboard.press('Control+v');
  check('paste translates formula', (await cell(2, 2))?.f === '=SUM(C1:C2)', JSON.stringify(await cell(2, 2)));

  // ---- fill handle: series 1,2 -> 1..5
  await page.evaluate(() => { const a = window.webExcel; const wb = a.doc.wb; wb.begin('t'); wb.setCell(wb.sheet, 0, 4, { v: 1 }); wb.setCell(wb.sheet, 1, 4, { v: 2 }); wb.commit(); a.selectRange({ r1: 0, c1: 4, r2: 1, c2: 4 }); });
  const e2 = await box(1, 4);
  await page.mouse.move(e2.l + e2.w - 1, e2.t + e2.h - 1);
  await page.mouse.down();
  const e5 = await box(4, 4);
  await page.mouse.move(e5.x, e5.y, { steps: 8 });
  await page.mouse.up();
  check('fill handle linear series', (await val(4, 4)) === 5, String(await val(4, 4)));

  // ---- Ctrl+D fill down
  await page.evaluate(() => window.webExcel.selectRange({ r1: 2, c1: 0, r2: 4, c2: 0 }));
  await page.keyboard.press('Control+d');
  check('Ctrl+D fills formula down', (await cell(4, 0)).f === '=SUM(A3:A4)', (await cell(4, 0)).f);
  await page.keyboard.press('Control+z');

  // ---- sheets
  await page.keyboard.press('Shift+F11');
  check('Shift+F11 new sheet', await page.evaluate(() => window.webExcel.doc.wb.sheets.length === 2));
  await page.keyboard.type('=Sheet1!A1*3'); await page.keyboard.press('Enter');
  check('cross-sheet formula', (await val(0, 0)) === 30);
  await page.evaluate(() => window.webExcel.doc.wb.renameSheet(0, '데이터 시트'));
  check('rename sheet updates formulas', await page.evaluate(() => window.webExcel.doc.wb.sheet.get(0, 0).f) === "='데이터 시트'!A1*3");
  await page.keyboard.press('Control+PageUp');
  check('Ctrl+PgUp previous sheet', await page.evaluate(() => window.webExcel.doc.wb.active === 0));

  // ---- open sample, filter, sort, chart
  await page.evaluate(() => window.webExcel.openFile(location.pathname && '/home/user/web-excel/samples/sample.xlsx'));
  await page.waitForFunction(() => window.webExcel.doc.wb.name === 'sample.xlsx');
  check('open xlsx: two docs', await page.evaluate(() => window.webExcel.docs.length === 2));
  check('xlsx formulas evaluated (XLOOKUP)', (await val(11, 1)) === 128, String(await val(11, 1)));
  check('xlsx chart loaded', await page.evaluate(() => window.webExcel.doc.wb.sheet.charts.length === 1));
  // filter menu on column B: keep only values > 150 via custom
  const fb = await box(0, 1);
  await page.mouse.click(fb.l + fb.w - 8, fb.y);
  await page.waitForSelector('.filter-menu');
  await page.click('.fm-custom summary');
  await page.selectOption('[data-cf=op0]', 'greaterThan');
  await page.fill('[data-cf=v0]', '150');
  await page.click('.filter-menu [data-act=ok]');
  const hidden = await page.evaluate(() => [...window.webExcel.doc.wb.sheet.filterHidden].sort());
  check('custom filter hides rows', JSON.stringify(hidden) === '[1,2]', JSON.stringify(hidden));
  check('SUBTOTAL respects filter', await page.evaluate(() => { const a = window.webExcel; return a.doc.wb.engine.evalFormula('=SUBTOTAL(9,B2:B7)', a.doc.wb.sheet, 20, 20); }) === 686);
  await page.screenshot({ path: path.join(OUT, 'filter.png') });
  await page.evaluate(() => window.webExcel.filters.clearAll());
  // sort descending by column C (inside autofilter range)
  await sel(1, 2);
  await page.evaluate(() => window.webExcel.filters.sortSelection(false));
  check('sort desc keeps header & moves rows', (await val(1, 2)) === 128 && (await val(0, 2)) === '제품B', `${await val(1, 2)}`);
  check('sorted row formula re-pointed', (await cell(1, 4)).f === '=SUM(B2:D2)', (await cell(1, 4)).f);
  check('totals unchanged after sort', (await val(7, 1)) === 951);
  // new chart from selection
  await page.evaluate(() => { const a = window.webExcel; a.selectRange({ r1: 0, c1: 0, r2: 6, c2: 2 }); a.quickChart('line'); });
  check('quick chart inserted', await page.evaluate(() => window.webExcel.doc.wb.sheet.charts.length === 2));
  await page.screenshot({ path: path.join(OUT, 'charts.png') });

  // ---- save as xlsx through Python, then reopen
  const outX = path.join(OUT, 'e2e_saved.xlsx');
  try { fs.unlinkSync(outX); } catch {}
  const ok1 = await page.evaluate((p) => window.webExcel.writeTo(p, 'xlsx', true), outX);
  check('save xlsx', ok1 && fs.existsSync(outX));
  await page.evaluate((p) => window.webExcel.openFile(p), outX);
  await page.waitForFunction(() => window.webExcel.doc.wb.name === 'e2e_saved.xlsx');
  check('reopen: charts persisted', await page.evaluate(() => window.webExcel.doc.wb.sheet.charts.length === 2));
  check('reopen: formulas persisted', (await cell(11, 1)).f === '=XLOOKUP("3월",A2:A7,C2:C7)', (await cell(11, 1)).f);
  check('reopen: autofilter persisted', await page.evaluate(() => !!window.webExcel.doc.wb.sheet.autoFilter));
  check('reopen: freeze persisted', await page.evaluate(() => JSON.stringify(window.webExcel.doc.wb.sheet.freeze) === '{"r":1,"c":1}'));
  // ---- .xls is read-only: Ctrl+S saves a new .xlsx next to it, the .xls stays as it was
  const inL = path.join(OUT, 'e2e_legacy.xls'), outL = path.join(OUT, 'e2e_legacy.xlsx');
  fs.copyFileSync(path.join(__dirname, '..', 'samples', 'sample.xls'), inL);
  try { fs.unlinkSync(outL); } catch {}
  const before = fs.readFileSync(inL);
  await page.evaluate((p) => window.webExcel.openFile(p), inL);
  await page.waitForFunction(() => window.webExcel.doc.wb.name === 'e2e_legacy.xls');
  check('xls opens', (await val(1, 0)) === '사과', String(await val(1, 0)));
  await page.evaluate(() => { const a = window.webExcel; a.select(5, 0); a.focusGrid(); });
  await page.keyboard.type('추가'); await page.keyboard.press('Enter');
  check('xls saved as xlsx next to it', await page.evaluate(() => window.webExcel.save()) && fs.existsSync(outL));
  check('original xls untouched', Buffer.compare(before, fs.readFileSync(inL)) === 0);
  check('document now points to the xlsx', await page.evaluate(() => window.webExcel.doc.wb.name) === 'e2e_legacy.xlsx');
  // open the .xls again: saving it now would replace the e2e_legacy.xlsx just written -> the prompt names it
  await page.evaluate((p) => window.webExcel.openFile(p), inL);
  await page.waitForFunction(() => window.webExcel.doc.wb.name === 'e2e_legacy.xls');
  const saving = page.evaluate(() => window.webExcel.save());
  await page.waitForSelector('.modal-overlay .modal-body', { timeout: 5000 });
  const prompt = await page.evaluate(() => document.querySelector('.modal-overlay .modal-body').textContent);
  check('overwrite prompt names the xlsx that would be replaced', /xlsx로 저장됩니다/.test(prompt) && prompt.includes("'e2e_legacy.xlsx'"), prompt);
  await page.click('.modal-overlay .modal-btns [data-act=no]');
  check('declining keeps the existing xlsx', (await saving) === false);

  // ---- multi-window sync
  await page.evaluate(() => window.webExcel.activateDoc(0));
  const [popup] = await Promise.all([ctx.waitForEvent('page'), page.evaluate(() => window.webExcel.newWindow())]);
  await popup.waitForFunction(() => window.webExcel?.doc?.wb);
  await popup.waitForTimeout(300);
  check('new window attached same doc', await popup.evaluate(() => window.webExcel.doc.wb.docId) === await page.evaluate(() => window.webExcel.doc.wb.docId));
  await popup.evaluate(() => { const a = window.webExcel; a.activateSheet(0); a.select(9, 0); });
  await popup.keyboard.type('창에서 입력'); await popup.keyboard.press('Enter');
  await page.waitForTimeout(400);
  check('edit in window B appears in window A', (await val(9, 0, 0)) === '창에서 입력', String(await val(9, 0, 0)));
  await popup.close();

  // ---- a table that does not start at A1 (title, blank rows, note, totals row)
  await page.evaluate((p) => window.webExcel.openFile(p), path.join(__dirname, '..', 'samples', 'offset.xlsx'));
  await page.waitForFunction(() => window.webExcel.doc.wb.name === 'offset.xlsx');
  await sel(6, 3);
  await page.keyboard.press('Control+Shift+L');
  check('UI filter from a cell in an offset table excludes title and totals', await page.evaluate(() => JSON.stringify(window.webExcel.doc.wb.sheet.autoFilter.range)) === '{"r1":4,"c1":2,"r2":11,"c2":5}');
  await sel(6, 3);
  await page.evaluate(() => window.webExcel.filters.sortSelection(false));
  check('UI sort keeps the totals row and its SUM', (await val(12, 2)) === '합계' && (await cell(12, 3)).f === '=SUM(D6:D12)' && (await val(5, 2)) === '대구');
  await sel(6, 3);
  await page.keyboard.press('Control+a');
  check('UI Ctrl+A selects the whole current region (with totals)', await page.evaluate(() => JSON.stringify(window.webExcel.view().sel)) === '{"r1":4,"c1":2,"r2":12,"c2":5}');

  await page.evaluate(() => window.webExcel.activateDoc(1));
  await page.screenshot({ path: path.join(OUT, 'final.png') });
  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  console.log(failures ? `\n${failures} FAILED` : '\nALL PASSED');
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
