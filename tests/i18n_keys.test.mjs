// Static checks of the UI text: every t('…') key used in the web UI has a Korean entry, the Korean
// dictionary has no entries nothing uses, and no Korean is left in the UI source outside the
// dictionary and the known data (fill lists, font names, number-format tokens, total labels).
//   node tests/i18n_keys.test.mjs
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const WEB = path.join(ROOT, '.agents/skills/monoonom/assets/web');
const KO = (await import(path.join(WEB, 'js/i18n/ko/index.js'))).default;
let fails = 0;
const check = (name, ok, extra = '') => { console.log(`${ok ? 'PASS' : 'FAIL'} ${name}${extra ? ' — ' + extra : ''}`); if (!ok) fails++; };

const files = fs.readdirSync(path.join(WEB, 'js')).filter((f) => f.endsWith('.js')).map((f) => path.join(WEB, 'js', f));
const src = Object.fromEntries(files.map((f) => [path.basename(f), fs.readFileSync(f, 'utf8')]));
const html = fs.readFileSync(path.join(WEB, 'index.html'), 'utf8');

// t('literal') / t("literal") / t(`literal without ${}`)
const used = new Map();
for (const [f, s] of Object.entries(src)) {
  for (const m of s.matchAll(/\bt\(\s*(['"`])((?:\\.|(?!\1)[^\\])*?)\1\s*[,)]/g)) {
    if (m[1] === '`' && m[2].includes('${')) continue;
    const k = m[2].replace(/\\(['"`\\])/g, '$1').replace(/\\n/g, '\n');
    if (!used.has(k)) used.set(k, f);
  }
}
const missing = [...used].filter(([k]) => !(k in KO.text)).map(([k, f]) => `${f}: ${k}`);
check(`every t() key has Korean text (${used.size} keys)`, missing.length === 0, missing.slice(0, 15).join(' | '));

// keys used elsewhere: index.html text (translateDom), data-i18n, labels rendered with t(variable)
const loose = Object.values(src).join('\n') + html;
const unused = Object.keys(KO.text).filter((k) => !used.has(k) && !loose.includes(k) && !loose.includes(JSON.stringify(k).slice(1, -1)));
check('no unused Korean entries', unused.length === 0, unused.slice(0, 15).join(' | '));

const placeholders = (s) => [...s.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort().join();
const bad = Object.entries(KO.text).filter(([en, ko]) => placeholders(en) !== placeholders(ko)).map(([en]) => en);
check('Korean entries keep the same {placeholders}', bad.length === 0, bad.slice(0, 10).join(' | '));

// Korean outside the dictionary: only known data
const HANGUL = /[가-힣]/;
const DATA = [
  /'일', '월', '화'|'일요일'|'1월', '2월'|'1분기'|'갑', '을'/, // Korean Excel's fill lists
  /FONT_CANDIDATES|FONT_FALLBACK|맑은 고딕|나눔고딕/, // font family names
  /mmmmmmmmmmlli가나다/, // font detection sample text
  /KO_DAYS|오전|오후|"년"|"월"|"일"|년|월\b/, // number-format tokens and Korean date/time input
  /TOTAL_LABEL|합\\s\*계/, // total-row labels
];
const stray = [];
for (const [f, s] of Object.entries(src)) {
  s.split('\n').forEach((line, i) => { if (HANGUL.test(line) && !DATA.some((re) => re.test(line))) stray.push(`${f}:${i + 1}: ${line.trim().slice(0, 90)}`); });
}
html.split('\n').forEach((line, i) => { if (HANGUL.test(line)) stray.push(`index.html:${i + 1}`); });
check('no Korean UI text outside js/i18n/ko/ (data excepted)', stray.length === 0, stray.slice(0, 12).join('\n  '));

console.log(fails ? `${fails} FAILED` : 'ALL PASSED');
process.exit(fails ? 1 : 0);
