// Row/column pixel geometry with sparse custom sizes and hidden rows/cols.
import { MAX_COLS, MAX_ROWS } from './address.js';

class Axis {
  constructor(def, sizes, hidden, max) {
    this.def = def;
    this.max = max;
    const map = new Map(sizes);
    for (const h of hidden) map.set(h, 0);
    this.idx = [...map.keys()].sort((a, b) => a - b);
    this.delta = this.idx.map((i) => map.get(i) - def);
    this.prefix = new Float64Array(this.idx.length + 1);
    for (let i = 0; i < this.idx.length; i++) this.prefix[i + 1] = this.prefix[i] + this.delta[i];
    this.sizeMap = map;
  }
  size(i) { const v = this.sizeMap.get(i); return v === undefined ? this.def : v; }
  /** count of overrides with index < i */
  _lower(i) {
    let lo = 0, hi = this.idx.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (this.idx[m] < i) lo = m + 1; else hi = m; }
    return lo;
  }
  pos(i) { return i * this.def + this.prefix[this._lower(i)]; }
  /** index containing position p (>=0) */
  at(p) {
    if (p <= 0) return 0;
    let lo = 0, hi = this.max - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.pos(m) <= p) lo = m; else hi = m - 1;
    }
    // skip zero-size (hidden) entries
    while (lo < this.max - 1 && this.size(lo) === 0) lo++;
    return lo;
  }
}

const cache = new WeakMap();
export function geom(sheet) {
  let g = cache.get(sheet);
  if (!g || g.v !== sheet._geomVersion) {
    g = {
      v: sheet._geomVersion,
      cols: new Axis(sheet.defaultColW, sheet.colW, sheet.hiddenCols, MAX_COLS),
      rows: new Axis(sheet.defaultRowH, sheet.rowH, new Set([...sheet.hiddenRows, ...sheet.filterHidden]), MAX_ROWS),
    };
    cache.set(sheet, g);
  }
  return g;
}
export function invalidate(sheet) { sheet._geomVersion++; }
