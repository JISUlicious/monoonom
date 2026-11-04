// Preview start-up and stream reconnects vs. documents the agent opens/closes at the same moment.
// Timing is made deterministic: the page's first /api/docs answer is held back (stale list), and the
// event stream is closed by hand while the agent works. Needs: python3 run.py --port 8765 --root . &
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XL = path.join(ROOT, '.agents', 'skills', 'web-excel', 'scripts', 'xl.py');
const OUT = path.join(ROOT, 'tests', 'out');
const BASE = 'http://localhost:8765';
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };
const xl = (...a) => execFileSync('python3', [XL, ...a, '--root', ROOT], { encoding: 'utf8' });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const closeAll = () => { for (const d of JSON.parse(xl('docs', '--json'))) xl('close', '--doc', d.docId, '--force'); };
const names = (page) => page.evaluate(() => window.webExcel.docs.map((d) => d.wb.name));
const active = (page) => page.evaluate(() => window.webExcel.doc.wb.name);

/** Load the page with its first /api/docs answer taken now but delivered after `during()` ran. */
async function loadWithStaleList(browser, url, during) {
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  let held = false;
  await page.route('**/api/docs', async (route) => {
    if (held || route.request().method() !== 'GET') return route.continue();
    held = true;
    const res = await route.fetch();       // the list as it is now
    await page.waitForFunction(() => window.webExcel?.stream?.ready, null, { timeout: 5000 }); // stream connected
    await during();                        // the agent opens documents meanwhile -> 'opened' events
    await sleep(300);
    await route.fulfill({ response: res }); // start() now sees the stale list
  });
  await page.goto(BASE + url);
  await page.waitForFunction(() => window.webExcel?.started, null, { timeout: 10000 });
  await sleep(500);
  return { page, errors };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  for (const n of ['a', 'b', 'c', 'd']) fs.copyFileSync(path.join(ROOT, 'samples', 'sample.xlsx'), path.join(OUT, `start-${n}.xlsx`));
  const browser = await chromium.launch();
  closeAll();

  // 1) server has no documents: start() creates Book1; the agent opens a file meanwhile
  let { page, errors } = await loadWithStaleList(browser, '/', async () => { xl('open', 'tests/out/start-a.xlsx'); });
  let n = await names(page);
  check('document opened during start-up gets a tab', n.includes('start-a.xlsx'), JSON.stringify(n));
  check('the blank Book1 made during start-up is replaced', n.length === 1, JSON.stringify(n));
  check('…and the opened document is shown', (await active(page)) === 'start-a.xlsx');
  check('no page errors (1)', errors.length === 0, errors.join('; '));
  await page.close();

  // 2) ?doc=A keeps focus on A; a document opened during start-up is added in the background
  const a = JSON.parse(xl('docs', '--json')).find((d) => d.name === 'start-a.xlsx').docId;
  xl('open', 'tests/out/start-b.xlsx');
  ({ page, errors } = await loadWithStaleList(browser, `/?doc=${a}`, async () => { xl('open', 'tests/out/start-c.xlsx'); }));
  n = await names(page);
  check('?doc window also gets the late document', n.includes('start-c.xlsx'), JSON.stringify(n));
  check('?doc=A stays on A', (await active(page)) === 'start-a.xlsx', await active(page));

  // 3) one listed document closed before it loads does not stop the others
  //    (close B between the stale list and its snapshot: the list still names it)
  await page.close();
  ({ page, errors } = await loadWithStaleList(browser, '/', async () => { xl('close', '--doc', 'start-b.xlsx', '--force'); xl('open', 'tests/out/start-d.xlsx'); }));
  n = await names(page);
  check('a document closed before loading is skipped, the rest attach', !n.includes('start-b.xlsx') && ['start-a.xlsx', 'start-c.xlsx', 'start-d.xlsx'].every((x) => n.includes(x)), JSON.stringify(n));
  check('no page errors (3)', errors.length === 0, errors.join('; '));

  // 4) stream down while the agent closes one document and opens another: reconnect catches up
  await page.evaluate(() => window.webExcel.stream.es.close());
  xl('close', '--doc', 'start-c.xlsx', '--force');
  xl('open', 'tests/out/start-b.xlsx');
  await page.evaluate(() => window.webExcel.stream.open());
  await page.waitForFunction(() => window.webExcel.docs.some((d) => d.wb.name === 'start-b.xlsx') && !window.webExcel.docs.some((d) => d.wb.name === 'start-c.xlsx'), null, { timeout: 5000 }).catch(() => {});
  n = await names(page);
  check('reconnect: document opened while down is added', n.includes('start-b.xlsx'), JSON.stringify(n));
  check('reconnect: document closed while down loses its tab', !n.includes('start-c.xlsx'), JSON.stringify(n));
  check('no page errors (4)', errors.length === 0, errors.join('; '));

  await browser.close();
  closeAll();
  console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
  process.exit(fails ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
