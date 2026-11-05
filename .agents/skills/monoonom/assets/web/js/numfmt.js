// Excel number-format renderer (subset of ECMA-376 Part 1 §18.8.31 numFmt codes).
// Supports sections (pos;neg;zero;text), conditions/colors, literals, thousands/scaling,
// decimals, percent, scientific, dates/times incl. [h], AM/PM and Korean 오전/오후.

const cache = new Map();
const DAY_MS = 86400000;
const KO_DAYS = ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'];
const EN_DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const EN_MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December'];

// Built-in format ids that openpyxl may surface as strings already; keep a few aliases.
export const FORMAT_PRESETS = [
  { label: '일반', code: 'General' },
  { label: '숫자', code: '0.00' },
  { label: '숫자(쉼표)', code: '#,##0' },
  { label: '통화(₩)', code: '"₩"#,##0' },
  { label: '통화($)', code: '"$"#,##0.00' },
  { label: '회계', code: '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-' },
  { label: '간단한 날짜', code: 'yyyy-mm-dd' },
  { label: '자세한 날짜', code: 'yyyy"년" m"월" d"일" dddd' },
  { label: '시간', code: 'h:mm:ss' },
  { label: '백분율', code: '0%' },
  { label: '백분율(소수)', code: '0.00%' },
  { label: '분수', code: '# ?/?' },
  { label: '지수', code: '0.00E+00' },
  { label: '텍스트', code: '@' },
];

/** Excel serial (1900 system, incl. the Lotus 1900-02-29 bug) -> UTC Date parts */
export function serialToParts(serial) {
  let days = Math.floor(serial);
  let frac = serial - days;
  let ms = Math.round(frac * DAY_MS);
  if (ms >= DAY_MS) { days += 1; ms -= DAY_MS; }
  let y, m, d, wd;
  if (days === 60) { y = 1900; m = 2; d = 29; wd = 3; }
  else {
    const base = days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    const dt = new Date(base + days * DAY_MS);
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate(); wd = dt.getUTCDay();
  }
  const secs = Math.floor(ms / 1000);
  return { y, m, d, wd, H: Math.floor(secs / 3600), M: Math.floor(secs / 60) % 60, S: secs % 60,
    ms: ms % 1000, totalH: days * 24 + Math.floor(secs / 3600), totalM: days * 1440 + Math.floor(secs / 60),
    totalS: days * 86400 + secs };
}

export function dateToSerial(y, m, d, H = 0, M = 0, S = 0) {
  // Excel DATE semantics: months/days overflow naturally
  const t = Date.UTC(y, m - 1, d, H, M, S);
  let serial = (t - Date.UTC(1899, 11, 30)) / DAY_MS;
  if (serial < 61) serial -= 1; // before 1900-03-01 the bug shifts by one
  if (serial < 60 && y === 1900 && m === 2 && d === 29) serial = 60;
  return serial;
}

function splitSections(code) {
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') q = !q;
    if (ch === '\\' && !q) { cur += ch + (code[i + 1] ?? ''); i++; continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function compileSection(src) {
  const sec = { color: null, cond: null, tokens: [], isDate: false, isText: false, percent: 0,
    hasNumber: false, src };
  // brackets
  let s = src.replace(/\[([^\]]*)\]/g, (all, inner) => {
    if (/^(black|white|red|green|blue|yellow|magenta|cyan)$/i.test(inner)) { sec.color = inner.toLowerCase(); return ''; }
    if (/^color\s*\d+$/i.test(inner)) return '';
    const cm = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner);
    if (cm) { sec.cond = { op: cm[1], v: parseFloat(cm[2]) }; return ''; }
    if (/^[hms]+$/i.test(inner)) return `\u0001${inner.toLowerCase()}\u0001`;
    const cur = /^\$([^-\]]*)(?:-[0-9A-Fa-f]+)?$/.exec(inner);
    if (cur) return cur[1] ? `"${cur[1]}"` : '';
    return '';
  });
  // tokenize
  const toks = [];
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      let j = i + 1, lit = '';
      while (j < s.length && s[j] !== '"') lit += s[j++];
      toks.push({ t: 'lit', v: lit }); i = j; continue;
    }
    if (ch === '\\') { toks.push({ t: 'lit', v: s[i + 1] ?? '' }); i++; continue; }
    if (ch === '_') { toks.push({ t: 'lit', v: ' ' }); i++; continue; }
    if (ch === '*') { i++; continue; }
    if (ch === '\u0001') {
      const j = s.indexOf('\u0001', i + 1);
      toks.push({ t: 'elapsed', v: s.slice(i + 1, j) }); sec.isDate = true; i = j; continue;
    }
    if (ch === '@') { toks.push({ t: 'text' }); sec.isText = true; continue; }
    const rest = s.slice(i);
    let m;
    if ((m = /^(AM\/PM|A\/P|오전\/오후)/i.exec(rest))) { toks.push({ t: 'ampm', v: m[1] }); sec.isDate = true; i += m[1].length - 1; continue; }
    if ((m = /^(yyyy|yy|e|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|aaaa|aaa)/i.exec(rest)) && !/^e[+-]/i.test(rest)) {
      toks.push({ t: 'date', v: m[1].toLowerCase() }); sec.isDate = true; i += m[1].length - 1; continue;
    }
    if ((m = /^\.0+/.exec(rest)) && toks.length && toks[toks.length - 1].t === 'date' && /^s/.test(toks[toks.length - 1].v)) {
      toks.push({ t: 'subsec', v: m[0].length - 1 }); i += m[0].length - 1; continue;
    }
    if ((m = /^(\d+|[#?]+)\s*\/\s*(\d+|[#?]+)/.exec(rest))) { toks.push({ t: 'frac', v: m[0] }); sec.hasNumber = true; i += m[0].length - 1; continue; }
    if ((m = /^[0#?,.]+(?:[eE][+-][0#]+)?/.exec(rest)) && /[0#?]/.test(m[0])) {
      toks.push({ t: 'num', v: m[0] }); sec.hasNumber = true; i += m[0].length - 1; continue;
    }
    if (ch === '%') { sec.percent++; toks.push({ t: 'lit', v: '%' }); continue; }
    toks.push({ t: 'lit', v: ch });
  }
  // m after h / before s means minutes
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t !== 'date' || !/^m{1,2}$/.test(t.v)) continue;
    const prev = toks.slice(0, i).reverse().find((x) => x.t === 'date' || x.t === 'elapsed');
    const next = toks.slice(i + 1).find((x) => x.t === 'date' || x.t === 'elapsed');
    if ((prev && /^h|^\[?h/.test(prev.v)) || (next && /^s/.test(next.v))) t.v = t.v === 'm' ? 'M' : 'MM';
  }
  sec.tokens = toks;
  sec.ampm = toks.some((t) => t.t === 'ampm');
  return sec;
}

function compile(code) {
  let c = cache.get(code);
  if (!c) {
    const parts = splitSections(code);
    c = parts.map(compileSection);
    cache.set(code, c);
  }
  return c;
}

const pad = (n, w) => String(n).padStart(w, '0');

function fmtNumberPattern(abs, pat) {
  // pat like "#,##0.00" or "0.00E+00"
  const em = /^([^eE]*)[eE]([+-])([0#]+)$/.exec(pat);
  if (em) {
    const [, mant, sign, expPat] = em;
    const intPart = mant.split('.')[0];
    const step = intPart.replace(/[^0#?]/g, '').length || 1; // engineering grouping
    let exp = abs === 0 ? 0 : Math.floor(Math.log10(abs));
    if (step > 1) exp = Math.floor(exp / step) * step;
    const m = abs === 0 ? 0 : abs / Math.pow(10, exp);
    const ms = fmtNumberPattern(m, mant);
    const es = pad(Math.abs(exp), expPat.length);
    return `${ms}E${exp < 0 ? '-' : (sign === '+' ? '+' : '')}${es}`;
  }
  let p = pat;
  let scale = 0;
  while (/,$/.test(p.replace(/\.[0#?]*$/, '')) && /[0#?],+(\.|$)/.test(p)) {
    p = p.replace(/,(?=\.|$)/, ''); scale++;
    if (scale > 4) break;
  }
  const v = abs / Math.pow(1000, scale);
  const [ip, dp = ''] = p.split('.');
  const decimals = dp.replace(/[^0#?]/g, '');
  const hasDot = p.includes('.');
  let fixed = v.toFixed(decimals.length);
  let [is, ds = ''] = fixed.split('.');
  // trim optional decimals (#)
  let minDec = decimals.replace(/[#?]+$/, '').length;
  if (decimals.length) {
    let k = ds.length;
    while (k > minDec && ds[k - 1] === '0' && decimals[k - 1] !== '0') k--;
    ds = ds.slice(0, k) + (decimals.slice(k).includes('?') ? ' '.repeat(decimals.length - k) : '');
  }
  const intDigits = ip.replace(/[^0#?]/g, '');
  const minInt = (intDigits.match(/0/g) || []).length;
  if (is === '0' && minInt === 0) is = '';
  is = is.padStart(minInt, '0');
  if (ip.includes(',')) is = is.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  // literal chars embedded in int pattern (rare), ignore
  return is + (hasDot && (ds.length || decimals.length === 0) ? (ds.length ? '.' + ds : '.') : (ds.length ? '.' + ds : ''));
}

function gcd(a, b) { return b ? gcd(b, a % b) : a; }

function fmtFraction(abs, pat) {
  const m = /^(\d+|[#?]+)\s*\/\s*(\d+|[#?]+)$/.exec(pat);
  const den = m && /^\d+$/.test(m[2]) ? parseInt(m[2], 10) : null;
  const maxDen = den ? den : Math.pow(10, m ? m[2].length : 1) - 1;
  const ip = Math.floor(abs);
  let f = abs - ip;
  let bestN = 0, bestD = 1;
  if (den) { bestN = Math.round(f * den); bestD = den; }
  else {
    let err = 1;
    for (let d = 1; d <= maxDen; d++) {
      const n = Math.round(f * d);
      const e = Math.abs(f - n / d);
      if (e < err - 1e-12) { err = e; bestN = n; bestD = d; }
    }
    const g = gcd(bestN, bestD) || 1; bestN /= g; bestD /= g;
  }
  return { ip, n: bestN, d: bestD };
}

function render(sec, value, showMinus) {
  const toks = sec.tokens;
  if (sec.isDate && typeof value === 'number') {
    const p = serialToParts(value + (toks.some((t) => t.t === 'subsec') ? 0 : 0.5 / 86400));
    let out = '';
    const H12 = ((p.H + 11) % 12) + 1;
    for (const t of toks) {
      switch (t.t) {
        case 'lit': out += t.v; break;
        case 'ampm': out += /오전/.test(t.v) ? (p.H < 12 ? '오전' : '오후') : t.v.length > 3 ? (p.H < 12 ? 'AM' : 'PM') : (p.H < 12 ? 'A' : 'P'); break;
        case 'elapsed': out += t.v[0] === 'h' ? pad(p.totalH, t.v.length) : t.v[0] === 'm' ? pad(p.totalM, t.v.length) : pad(p.totalS, t.v.length); break;
        case 'subsec': out += '.' + pad(Math.floor(p.ms / Math.pow(10, 3 - t.v)), t.v); break;
        case 'date': {
          const v = t.v;
          if (v === 'yyyy' || v === 'e') out += p.y;
          else if (v === 'yy') out += pad(p.y % 100, 2);
          else if (v === 'mmmmm') out += EN_MON[p.m - 1][0];
          else if (v === 'mmmm') out += EN_MON[p.m - 1];
          else if (v === 'mmm') out += EN_MON[p.m - 1].slice(0, 3);
          else if (v === 'mm') out += pad(p.m, 2);
          else if (v === 'm') out += p.m;
          else if (v === 'MM') out += pad(p.M, 2);
          else if (v === 'M') out += p.M;
          else if (v === 'dddd') out += EN_DAYS[p.wd];
          else if (v === 'ddd') out += EN_DAYS[p.wd].slice(0, 3);
          else if (v === 'aaaa') out += KO_DAYS[p.wd];
          else if (v === 'aaa') out += KO_DAYS[p.wd][0];
          else if (v === 'dd') out += pad(p.d, 2);
          else if (v === 'd') out += p.d;
          else if (v === 'hh') out += pad(sec.ampm ? H12 : p.H, 2);
          else if (v === 'h') out += sec.ampm ? H12 : p.H;
          else if (v === 'ss') out += pad(p.S, 2);
          else if (v === 's') out += p.S;
          break;
        }
        default: break;
      }
    }
    return out;
  }
  if (typeof value !== 'number') {
    let out = '';
    for (const t of toks) {
      if (t.t === 'text') out += value; else if (t.t === 'lit') out += t.v;
    }
    return sec.isText ? out : String(value);
  }
  const v = Math.abs(value) * Math.pow(100, sec.percent);
  const fracTok = toks.find((t) => t.t === 'frac');
  let out = '';
  if (fracTok) {
    const fr = fmtFraction(v, fracTok.v.replace(/\s/g, ''));
    const mixed = toks.indexOf(fracTok) > toks.findIndex((t) => t.t === 'num') && toks.some((t) => t.t === 'num');
    let didNum = false;
    for (const t of toks) {
      if (t.t === 'lit') out += t.v;
      else if (t.t === 'num' && !didNum && mixed) { out += fr.ip ? fr.ip : (fr.n ? '' : '0'); didNum = true; }
      else if (t.t === 'frac') out += mixed ? (fr.n ? `${fr.n}/${fr.d}` : '') : `${fr.ip * fr.d + fr.n}/${fr.d}`;
    }
    out = out.trim();
  } else {
    let numDone = false;
    for (const t of toks) {
      if (t.t === 'lit') out += t.v;
      else if (t.t === 'text') out += String(value);
      else if (t.t === 'num' && !numDone) { out += fmtNumberPattern(v, t.v); numDone = true; }
    }
  }
  return (showMinus && /[1-9]/.test(out) ? '-' : '') + out;
}

function generalNumber(v) {
  if (!isFinite(v)) return '#NUM!';
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) {
    const [mant, ex] = v.toExponential(5).split('e');
    const e = parseInt(ex, 10);
    return mant.replace(/\.?0+$/, '') + 'E' + (e < 0 ? '-' : '+') + pad(Math.abs(e), 2);
  }
  let s = String(parseFloat(v.toPrecision(10)));
  if (s.replace(/[-.]/g, '').length > 10) {
    const intLen = Math.floor(Math.log10(a)) + 1;
    s = String(parseFloat(v.toFixed(Math.max(0, 10 - Math.max(intLen, 0)))));
  }
  return s;
}

/**
 * Format a value. Returns {text, color} where color is a CSS color or null.
 */
export function formatValue(value, code) {
  if (value === null || value === undefined || value === '') return { text: '', color: null };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', color: null };
  if (typeof value === 'object' && value.err) return { text: value.err, color: null };
  if (!code || code === 'General' || code === 'general') {
    return { text: typeof value === 'number' ? generalNumber(value) : String(value), color: null };
  }
  const secs = compile(code);
  if (typeof value !== 'number') {
    const ts = secs.length >= 4 ? secs[3] : secs.find((x) => x.isText);
    return ts ? { text: render(ts, value, false), color: ts.color } : { text: String(value), color: null };
  }
  const n = secs.length;
  let sec, showMinus;
  if (secs.some((x) => x.cond)) {
    const test = (x) => {
      if (!x.cond) return true;
      const { op, v } = x.cond;
      return op === '<' ? value < v : op === '>' ? value > v : op === '=' ? value === v
        : op === '<=' ? value <= v : op === '>=' ? value >= v : value !== v;
    };
    sec = secs.slice(0, 2).find(test) || secs[2] || secs[0];
    showMinus = value < 0 && !(sec.cond && sec.cond.v <= 0 && /</.test(sec.cond.op));
  } else if (value > 0 || n === 1) { sec = secs[0]; showMinus = value < 0; }
  else if (value < 0) { sec = secs[1]; showMinus = false; }
  else { sec = n >= 3 ? secs[2] : secs[0]; showMinus = false; }
  if (sec.isText && !sec.hasNumber) return { text: generalNumber(value), color: sec.color };
  if (sec.isDate && (value < 0 || value > 2958465)) return { text: '#'.repeat(8), color: null };
  if (!sec.isDate && !sec.hasNumber && !sec.tokens.length) return { text: '', color: null };
  return { text: render(sec, value, showMinus), color: sec.color };
}

export function isDateFormat(code) {
  if (!code || code === 'General') return false;
  return compile(code).some((s) => s.isDate);
}

export const COLOR_NAMES = { black: '#000', white: '#fff', red: '#f00', green: '#008000', blue: '#00f',
  yellow: '#ff0', magenta: '#f0f', cyan: '#0ff' };

/** Parse typed user input into a value and an implied number format (like Excel). */
export function parseInput(text) {
  const t = text.trim();
  if (t === '') return { v: null };
  if (/^(true|false)$/i.test(t)) return { v: t.toUpperCase() === 'TRUE' };
  if (/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/i.test(t)) return { v: t.toUpperCase(), t: 'e' };
  let m;
  if ((m = /^([-+]?)(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$/.exec(t)) && (m[2] || m[4])) {
    const v = parseFloat(t.replace(/,/g, ''));
    return { v, nf: m[3] ? (m[4] ? '#,##0.00' : '#,##0') : null };
  }
  if ((m = /^([-+]?[\d,]*\.?\d+)\s*%$/.exec(t))) return { v: parseFloat(m[1].replace(/,/g, '')) / 100, nf: /\./.test(m[1]) ? '0.00%' : '0%' };
  if ((m = /^([-+]?)([₩$])\s*([\d,]*\.?\d+)$/.exec(t))) {
    const v = parseFloat(m[3].replace(/,/g, '')) * (m[1] === '-' ? -1 : 1);
    return { v, nf: m[2] === '₩' ? '"₩"#,##0' : '"$"#,##0.00' };
  }
  if (/^[-+]?\d+(\.\d+)?[eE][-+]?\d+$/.test(t)) return { v: parseFloat(t), nf: '0.00E+00' };
  // dates: yyyy-mm-dd, yyyy/mm/dd, yyyy.mm.dd, m/d, with optional time
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\.?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    const [y, mo, d] = [+m[1], +m[2], +m[3]];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      const v = dateToSerial(y, mo, d, +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
      return { v, nf: m[4] ? 'yyyy-mm-dd h:mm' : 'yyyy-mm-dd' };
    }
  }
  if ((m = /^(\d{1,2})[/-](\d{1,2})$/.exec(t))) {
    const mo = +m[1], d = +m[2];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      return { v: dateToSerial(new Date().getFullYear(), mo, d), nf: 'm"월" d"일"' };
    }
  }
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|오전|오후)?$/i.exec(t))) {
    let h = +m[1];
    const ap = (m[4] || '').toUpperCase();
    if ((ap === 'PM' || ap === '오후') && h < 12) h += 12;
    if ((ap === 'AM' || ap === '오전') && h === 12) h = 0;
    return { v: (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, nf: m[3] ? 'h:mm:ss' : 'h:mm' };
  }
  return { v: text };
}
