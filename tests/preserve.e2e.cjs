// Browser side of "unsupported functions": warning on open, F2+Enter keeps the file's formula,
// editing one warns. Needs the dev server: python3 run.py --port 8765 --root . &
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'web-excel', 'scripts', 'xl.py');
const PYTEST = path.join(ROOT, 'tests', 'preserve_formulas.test.py');
const OUT = path.join(ROOT, 'tests', 'out');
const BASE = 'http://localhost:8765';
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ROOT], { encoding: 'utf8' });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const src = path.join(OUT, 'keep.xlsx');
  execFileSync('python3', [PYTEST, '--make', src]);
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1300, height: 800 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(BASE + '/');
  await page.waitForFunction(() => window.webExcel?.doc?.sync);
  xl('open', 'tests/out/keep.xlsx');
  await page.waitForFunction(() => window.webExcel.docs.some((d) => d.wb.name === 'keep.xlsx'), null, { timeout: 8000 });
  await page.evaluate(() => { const a = window.webExcel; a.activateDoc(a.docs.findIndex((d) => d.wb.name === 'keep.xlsx')); });
  await sleep(500);
  const toasts = await page.evaluate(() => [...document.querySelectorAll('#toasts .toast')].map((t) => t.textContent).join('\n'));
  check('warning shown when the file opens', /계산할 수 없는 함수/.test(toasts) && /LAMBDA/.test(toasts), toasts.slice(0, 120));
  const val = (r, c) => page.evaluate(([r, c]) => { const a = window.webExcel; const v = a.doc.wb.value(a.doc.wb.sheet, r, c); return v && v.err ? v.err : v; }, [r, c]);
  check('unsupported cells show the cached results', JSON.stringify([await val(0, 1), await val(1, 1), await val(2, 1)]) === '[6,"a",42]');

  // F2 + Enter without changes on B1, then a real edit on B2
  const ver = () => page.evaluate(() => window.webExcel.doc.sync.version);
  await page.evaluate(() => { const a = window.webExcel; a.select(0, 1); a.focusGrid(); });
  const v0 = await ver();
  await page.keyboard.press('F2'); await page.keyboard.press('Enter');
  await sleep(300);
  check('F2+Enter on a formula this parser cannot read leaves the editor', !(await page.evaluate(() => window.webExcel.editor.active)));
  check('… and changes nothing', (await ver()) === v0, `${v0} -> ${await ver()}`);
  await page.evaluate(() => { const a = window.webExcel; a.select(1, 1); a.focusGrid(); });
  await page.keyboard.type('=TEXTSPLIT("x;y",";")'); await page.keyboard.press('Enter');
  await sleep(300);
  const t2 = await page.evaluate(() => [...document.querySelectorAll('#toasts .toast')].map((t) => t.textContent).join('\n'));
  check('editing such a cell warns', /새 수식이 저장되고/.test(t2));
  check('re-entered B1 still shows the cached result', (await val(0, 1)) === 6, String(await val(0, 1)));
  await sleep(700);
  xl('save', '--doc', 'keep.xlsx', '--as', 'tests/out/keep-saved.xlsx', '--overwrite');
  const cells = JSON.parse(execFileSync('python3', [PYTEST, '--cells', src, path.join(OUT, 'keep-saved.xlsx')], { encoding: 'utf8' }));
  check('B1 (F2+Enter) written back unchanged', JSON.stringify(cells.new.B1) === JSON.stringify(cells.old.B1), JSON.stringify(cells.new.B1));
  check('B2 edited: new formula with _xlfn. prefix', cells.new.B2[0] === '_xlfn.TEXTSPLIT(&quot;x;y&quot;,&quot;;&quot;)' || cells.new.B2[0] === '_xlfn.TEXTSPLIT("x;y",";")', JSON.stringify(cells.new.B2));
  check('B3 (untouched VBA function) unchanged', JSON.stringify(cells.new.B3) === JSON.stringify(cells.old.B3));
  check('no page errors', errors.length === 0, errors.join('; '));
  xl('close', '--doc', 'keep-saved.xlsx', '--force');
  await browser.close();
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
