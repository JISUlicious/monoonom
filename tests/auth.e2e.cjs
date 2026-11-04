// Browser auth in the ways previews embed the page. Starts its own servers (ports 8771/8772).
//   node tests/auth.e2e.cjs
// The page URL is http://localhost:PORT/#t=TOKEN; the token goes in the X-WebExcel-Token header, so the
// page also works inside a cross-site iframe (VS Code Simple Browser, Tauri apps) where cookies are not sent.
const { chromium } = require('playwright');
const { execFileSync, spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'web-excel', 'scripts', 'xl.py');
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const xl = (ws, ...a) => execFileSync('python3', [XL, ...a, '--root', ws], { encoding: 'utf8' });
const state = (ws) => JSON.parse(fs.readFileSync(path.join(ws, '.web-excel', 'server.json'), 'utf8'));

function workspace(name) {
  const ws = fs.mkdtempSync(path.join(os.tmpdir(), `wx-auth-${name}-`));
  fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(ws, `${name}.xlsx`));
  return ws;
}
// a page on another site that embeds the preview URL in an iframe, like an app's webview does
function hostPage(port) {
  const srv = http.createServer((req, res) => {
    const src = decodeURIComponent((req.url.split('src=')[1] || ''));
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(`<!doctype html><iframe src="${src}" style="width:1200px;height:700px"></iframe>`);
  });
  return new Promise((r) => srv.listen(port, '127.0.0.1', () => r(srv)));
}
const framed = (page) => page.frames().find((f) => f.url().startsWith('http://localhost:'));
const docsIn = (f) => f.evaluate(() => (window.webExcel?.docs || []).map((d) => d.wb.name));
const toasts = (f) => f.evaluate(() => [...document.querySelectorAll('#toasts .toast')].map((t) => t.textContent).join(' | '));

(async () => {
  const wsA = workspace('alpha'), wsB = workspace('beta');
  xl(wsA, 'serve', '--port', '8771'); xl(wsA, 'open', 'alpha.xlsx');
  xl(wsB, 'serve', '--port', '8772'); xl(wsB, 'open', 'beta.xlsx');
  const host = await hostPage(8779);
  const browser = await chromium.launch();
  const errors = [];
  try {
    const urlA = state(wsA).url;
    check('serve prints the URL with #t=TOKEN', /^http:\/\/localhost:8771\/#t=.{16,}$/.test(urlA), urlA);
    check('status shows the same URL', xl(wsA, 'status').includes(urlA));

    // 1) cross-site iframe with the token: documents load, live sync works
    const ctx = await browser.newContext();
    let page = await ctx.newPage();
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(`http://127.0.0.1:8779/?src=${encodeURIComponent(urlA)}`);
    await sleep(2500);
    let f = framed(page);
    check('iframe on another site + #t: workbook opens', JSON.stringify(await docsIn(f)) === '["alpha.xlsx"]', JSON.stringify(await docsIn(f)));
    check('…token is removed from the address', !(await f.evaluate(() => location.hash)).includes('t='));
    xl(wsA, 'write', 'A20', '에이전트');
    await f.waitForFunction(() => window.webExcel.doc.wb.sheet.get(19, 0)?.v === '에이전트', null, { timeout: 5000 }).catch(() => {});
    check('…agent edits arrive over the event stream (ticket)', await f.evaluate(() => window.webExcel.doc.wb.sheet.get(19, 0)?.v === '에이전트'));
    await f.evaluate(() => { const a = window.webExcel; a.select(20, 0); a.focusGrid(); });
    await page.keyboard.type('사용자'); await page.keyboard.press('Enter');
    await sleep(600);
    check('…user edits reach the server', xl(wsA, 'read', 'A21').includes('사용자'));

    // 2) same iframe without the token: clear explanation instead of a silent failure
    page = await ctx.newPage();
    await page.goto(`http://127.0.0.1:8779/?src=${encodeURIComponent(state(wsA).base_url)}`);
    await sleep(2000);
    f = framed(page);
    const t = await toasts(f);
    check('iframe without #t: says the token is missing', /인증 정보가 없습니다/.test(t) && /#t=/.test(t), t.slice(0, 120));
    await page.close();

    // 3) two servers on one host: each tab keeps working (cookies are per host, not per port)
    const top = await ctx.newPage();
    await top.goto(state(wsA).base_url); await sleep(800);   // cookie only (no #t)
    const topB = await ctx.newPage();
    await topB.goto(state(wsB).base_url); await sleep(800);
    const st = (p) => p.evaluate(async () => (await fetch('/api/docs')).status);
    check('two servers: first tab still authorized after opening the second', (await st(top)) === 200 && (await st(topB)) === 200);

    // 4) server restarted while the preview is open: same token, the page keeps working
    xl(wsA, 'stop', '--force'); xl(wsA, 'serve', '--port', '8771'); xl(wsA, 'open', 'alpha.xlsx');
    check('restart keeps the workspace token', state(wsA).url === urlA);
    f = framed((await ctx.pages())[0]);
    await sleep(3000);
    // (the restart closed the documents; windows that lost their last tab make a blank one, so name the target)
    xl(wsA, 'write', 'B20', '재시작', '--doc', 'alpha.xlsx');
    const ok = await f.waitForFunction(() => window.webExcel.docs.some((d) => d.wb.sheet.get(19, 1)?.v === '재시작'), null, { timeout: 8000 }).then(() => true).catch(() => false);
    check('…the open iframe reconnects and syncs again (no reload)', ok);

    // 5) the log never contains the token or a stream ticket
    const log = fs.readFileSync(path.join(wsA, '.web-excel', 'server.log'), 'utf8');
    check('server log has no token or ticket', !log.includes(urlA.split('#t=')[1]) && !/ticket=(?!\*\*\*)/.test(log));
    check('no page errors', errors.length === 0, errors.join('; '));
  } finally {
    await browser.close(); host.close();
    for (const ws of [wsA, wsB]) { try { xl(ws, 'stop', '--force'); } catch {} fs.rmSync(ws, { recursive: true, force: true }); }
  }
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
