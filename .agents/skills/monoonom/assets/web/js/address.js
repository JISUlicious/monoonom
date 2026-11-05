// A1 reference utilities (0-based rows/cols internally).
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export const key = (r, c) => r * MAX_COLS + c;
export const keyRow = (k) => Math.floor(k / MAX_COLS);
export const keyCol = (k) => k % MAX_COLS;

export function colName(c) {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(name) {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const a1 = (r, c) => colName(c) + (r + 1);

export function parseA1(s) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim());
  if (!m) return null;
  return { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) };
}

/** "A1:B3" | "A1" | "A:C" | "2:5" -> {r1,c1,r2,c2} normalized */
export function parseRange(s) {
  s = s.replace(/\$/g, '').trim();
  const [p, q = p] = s.split(':');
  let r1, c1, r2, c2;
  if (/^[A-Za-z]+$/.test(p) && /^[A-Za-z]+$/.test(q)) {
    c1 = colIndex(p); c2 = colIndex(q); r1 = 0; r2 = MAX_ROWS - 1;
  } else if (/^\d+$/.test(p) && /^\d+$/.test(q)) {
    r1 = +p - 1; r2 = +q - 1; c1 = 0; c2 = MAX_COLS - 1;
  } else {
    const a = parseA1(p), b = parseA1(q);
    if (!a || !b) return null;
    r1 = a.r; c1 = a.c; r2 = b.r; c2 = b.c;
  }
  return normRange({ r1, c1, r2, c2 });
}

export function normRange(g) {
  return {
    r1: Math.min(g.r1, g.r2), c1: Math.min(g.c1, g.c2),
    r2: Math.max(g.r1, g.r2), c2: Math.max(g.c1, g.c2),
  };
}

export function rangeA1(g) {
  const a = a1(g.r1, g.c1), b = a1(g.r2, g.c2);
  return a === b ? a : `${a}:${b}`;
}

export function absRangeA1(g) {
  const f = (r, c) => `$${colName(c)}$${r + 1}`;
  return g.r1 === g.r2 && g.c1 === g.c2 ? f(g.r1, g.c1) : `${f(g.r1, g.c1)}:${f(g.r2, g.c2)}`;
}

export function quoteSheet(name) {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name : `'${name.replace(/'/g, "''")}'`;
}

/** "'Sheet 1'!$A$1:$B$2" -> {sheet, range} */
export function parseSheetRef(ref) {
  if (!ref) return null;
  ref = ref.trim().replace(/^=/, '');
  let sheet = null, rest = ref;
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref);
  if (m) { sheet = (m[1] ?? m[2]).replace(/''/g, "'"); rest = m[3]; }
  const range = parseRange(rest);
  return range ? { sheet, range } : null;
}

export const inRange = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
export const rangesIntersect = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
