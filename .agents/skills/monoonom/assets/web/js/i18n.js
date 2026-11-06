// UI language: English by default, Korean when the browser prefers Korean. `?lang=en|ko` in the page
// URL overrides it and is remembered in this browser (localStorage 'monoonom.lang').
//
// The English text is the key: t('Save') is 'Save' in English and the Korean dictionary's entry in
// Korean. All Korean UI text lives in i18n/ko/ (one file per module, merged in i18n/ko/index.js).
// Text that is user content (sheet names, cell values, file names) is never translated; mark such
// elements translate="no" so translateDom() skips them.
import KO from './i18n/ko/index.js';

function detect() {
  if (typeof location === 'undefined') return 'en'; // Node (calc worker, tests)
  let q = null;
  try { q = new URLSearchParams(location.search).get('lang'); } catch { /* ignore */ }
  if (q === 'en' || q === 'ko') {
    try { localStorage.setItem('monoonom.lang', q); } catch { /* storage blocked */ }
    return q;
  }
  try { const s = localStorage.getItem('monoonom.lang'); if (s === 'en' || s === 'ko') return s; } catch { /* ignore */ }
  return /^ko\b/i.test((typeof navigator !== 'undefined' && navigator.language) || '') ? 'ko' : 'en';
}

export const LANG = detect();
const DICT = LANG === 'ko' ? KO.text : null;

/** The UI-language text for an English string; {name} placeholders are filled from vars. */
export function t(s, vars) {
  const out = (DICT && DICT[s]) ?? s;
  return vars ? out.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m)) : out;
}

/** Number formats and the new-workbook font that follow the UI language (as Excel in that language). */
export const FORMATS = { currency: '"$"#,##0.00', currencySymbol: '$', currencyDecimals: 2, longDate: 'dddd, mmmm d, yyyy', monthDay: 'd-mmm', defaultFont: 'Calibri', ...(LANG === 'ko' ? KO.formats : {}) };

/** A per-language value: L({ en: 'yyyy-mm-dd', ko: … }). */
export const L = (o) => (LANG in o ? o[LANG] : o.en);

/** True when an English key has a translation (tests, fallbacks). */
export const has = (s) => !DICT || s in DICT;

// a text node may carry symbols around its words ("✂ Cut", "⊕ Insert ▾", "Σ AutoSum"): try the
// whole text, then without the symbols at its ends, then without a short leading icon token
function tr(text) {
  const s = text.trim();
  if (!s) return null;
  const cands = [s, s.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N})]+$/gu, '')];
  const m = /^\S{1,3}\s+(.+)$/su.exec(s);
  if (m) cands.push(m[1], m[1].replace(/[^\p{L}\p{N})]+$/u, ''));
  for (const c of cands) if (c && c in DICT) return text.replace(c, DICT[c]);
  return null;
}
const ATTRS = ['title', 'placeholder', 'aria-label'];

/** Translate the static text of a DOM subtree in place: text nodes and title / placeholder /
 * aria-label attributes whose text is a dictionary key, and elements with data-i18n="key" (their
 * text becomes that key's translation). Skips translate="no" subtrees. Returns root. */
export function translateDom(root) {
  if (!DICT || !root) return root;
  const walk = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.nodeType === 1 && n.getAttribute('translate') === 'no' ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  const visit = (n) => {
    if (n.nodeType === 3) {
      const v = tr(n.nodeValue);
      if (v != null) n.nodeValue = v;
      return;
    }
    for (const a of ATTRS) {
      const v = n.hasAttribute(a) ? tr(n.getAttribute(a)) : null;
      if (v != null) n.setAttribute(a, v);
    }
    const k = n.dataset?.i18n;
    if (k && k in DICT) n.textContent = DICT[k];
  };
  if (root.nodeType === 1 && root.getAttribute('translate') === 'no') return root;
  visit(root);
  for (let n = walk.nextNode(); n; n = walk.nextNode()) visit(n);
  return root;
}

/** Server messages (errors, warnings, notes) are English; show them in the UI language where a
 * pattern in the dictionary matches. */
export function serverText(msg) {
  if (!DICT || typeof msg !== 'string') return msg;
  for (const [re, ko] of KO.server) if (re.test(msg)) return msg.replace(re, ko);
  return msg;
}
