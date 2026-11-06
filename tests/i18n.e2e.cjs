// The preview's language: English by default, Korean for a Korean browser, ?lang= overrides (and is
// remembered). English pages show no Korean UI text; server messages and new-workbook defaults
// follow the language. The other browser tests run in Korean (tests/korean.cjs).
// Starts its own server (port 8874).   node tests/i18n.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-i18n-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const HANGUL = /[가-힣]/;

(async () => {
  xl('serve', '--port', '8874');
  const url = xl('url', '--bare').trim();
  const browser = await chromium.launch();
  const errors = [];
  try {
    // English (default locale en-US)
    const ctx = await browser.newContext({ locale: 'en-US', viewport: { width: 1300, height: 800 } });
    const p = await ctx.newPage();
    p.on('pageerror', (e) => errors.push(e.message));
    await p.goto(url);
    await p.waitForFunction(() => window.monoonom?.started); await sleep(300);
    const start = await p.evaluate(() => ({ lang: document.documentElement.lang, home: document.querySelector('[data-rtab=home]').textContent, text: document.body.innerText }));
    check('English by default: html lang, ribbon, start screen', start.lang === 'en' && start.home === 'Home' && /No file is open|agent opens a file/i.test(start.text), JSON.stringify(start).slice(0, 300));
    check('…no Korean anywhere on the start screen', !HANGUL.test(start.text), (start.text.match(/.{0,20}[가-힣].{0,20}/) || [''])[0]);
    await p.click('#empty-state [data-act=new]');
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    const doc = await p.evaluate(() => ({ name: window.monoonom.doc.wb.name, font: window.monoonom.doc.wb.defaultFont.name, ui: document.body.innerText, nf: [...document.querySelectorAll('#num-format option')].map((o) => o.value) }));
    check('new workbook from the English page: Book1.xlsx, Calibri', doc.name === 'Book1.xlsx' && doc.font === 'Calibri', `${doc.name} ${doc.font}`);
    check('…currency and long date formats in the English variants', doc.nf.includes('"$"#,##0.00') && doc.nf.includes('dddd, mmmm d, yyyy'), doc.nf.join(' | '));
    check('…no Korean in the UI with a workbook open', !HANGUL.test(doc.ui), (doc.ui.match(/.{0,20}[가-힣].{0,20}/) || [''])[0]);
    // context menu, a dialog and a ribbon dropdown
    await p.mouse.click(300, 300, { button: 'right' }); await sleep(200);
    const menu = await p.evaluate(() => document.querySelector('.ctx-menu, .menu')?.innerText || '');
    check('context menu in English', /Cut/.test(menu) && /Copy/.test(menu) && !HANGUL.test(menu), menu.slice(0, 200));
    await p.keyboard.press('Escape');
    await p.evaluate(() => window.monoonom.exec('formatCells')); await sleep(300);
    const dlg = await p.evaluate(() => document.querySelector('.modal, .dialog')?.innerText || '');
    check('Format Cells dialog in English', dlg.length > 20 && !HANGUL.test(dlg), dlg.slice(0, 200));
    await p.keyboard.press('Escape'); await sleep(150);
    await p.evaluate(() => window.monoonom.exec('help')); await sleep(300);
    const help = await p.evaluate(() => document.querySelector('.kbd-table')?.innerText || '');
    check('keyboard shortcuts table in English', /Undo/.test(help) && !HANGUL.test(help), help.slice(0, 120));
    await p.keyboard.press('Escape');
    const fn = await p.evaluate(async () => (await import('/js/fnhelp.js')).fnHelp('SUM'));
    check('function help in English', fn && !HANGUL.test(fn.join(' ')), JSON.stringify(fn));
    // agent's pivot from an English workbook: English labels by default
    // ?lang=ko switches to Korean and is remembered
    const ko = await ctx.newPage();
    await ko.goto(url.replace('/#', '/?lang=ko#'));
    await ko.waitForFunction(() => window.monoonom?.started); await sleep(300);
    const k = await ko.evaluate(async () => {
      const i18n = await import('/js/i18n.js');
      return { lang: document.documentElement.lang, home: document.querySelector('[data-rtab=home]').textContent, server: i18n.serverText('File not found: a.xlsx'), bold: document.querySelector('[data-cmd=bold]').textContent };
    });
    check('?lang=ko: Korean UI', k.lang === 'ko' && k.home === '홈' && k.bold === '가', JSON.stringify(k));
    check('…server messages shown in Korean', HANGUL.test(k.server) && k.server.includes('a.xlsx'), k.server);
    await ko.goto(url); await ko.waitForFunction(() => window.monoonom?.started); await sleep(200);
    check('…remembered without ?lang', await ko.evaluate(() => document.documentElement.lang) === 'ko');
    await ko.goto(url.replace('/#', '/?lang=en#')); await ko.waitForFunction(() => window.monoonom?.started);
    // Korean browser locale without any setting
    const kctx = await browser.newContext({ locale: 'ko-KR' });
    const kp = await kctx.newPage();
    await kp.goto(url); await kp.waitForFunction(() => window.monoonom?.started); await sleep(200);
    check('Korean browser (ko-KR): Korean UI', await kp.evaluate(() => document.querySelector('[data-rtab=home]').textContent) === '홈');
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close();
  }
  // agent side: English labels by default, Korean with --lang ko
  const out = xl('new', '--lang', 'ko');
  check('xl.py new --lang ko: Korean Excel default name', /통합 문서/.test(out), out.trim());
  try { xl('stop', '--force'); } catch { /* ignore */ }
  fs.rmSync(ws, { recursive: true, force: true });
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
