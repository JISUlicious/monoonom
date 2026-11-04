// Canvas-virtualized spreadsheet grid: rendering, hit-testing, selection, resizing, fill handle.
import { key, colName, MAX_COLS, MAX_ROWS, normRange, inRange } from './address.js';
import { formatValue, COLOR_NAMES } from './numfmt.js';
import { geom } from './geometry.js';

const HDR_H = 22;
const GRID_COLOR = '#e0e3e7';
const HDR_BG = '#f5f6f7';
const HDR_SEL_BG = '#dcefe4';
const ACCENT = '#217346';
const FONT_FALLBACK = '"Malgun Gothic", "맑은 고딕", "Apple SD Gothic Neo", "Noto Sans KR", Calibri, Arial, sans-serif';
const BORDER_W = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const BORDER_DASH = { dotted: [1, 2], hair: [1, 1], dashed: [4, 2], mediumDashed: [6, 3], dashDot: [5, 2, 1, 2], mediumDashDot: [6, 3, 2, 3], dashDotDot: [5, 2, 1, 2, 1, 2], mediumDashDotDot: [6, 2, 2, 2, 2, 2], slantDashDot: [6, 2, 2, 2] };

export const ptToPx = (pt) => (pt * 96) / 72;

export class Grid {
  constructor(app, host) {
    this.app = app;
    this.host = host;
    host.innerHTML = '';
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroller';
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    this.chartLayer = document.createElement('div');
    this.chartLayer.className = 'chart-layer';
    this.scroller.append(this.spacer, this.chartLayer);
    host.append(this.canvas, this.scroller);
    this.ctx = this.canvas.getContext('2d');
    this.refHighlights = [];
    this.fillPreview = null;
    this.measureCache = new Map();
    this.marchOffset = 0;
    this._raf = 0;
    this.bind();
    new ResizeObserver(() => this.render()).observe(host);
  }

  // ------------------------------------------------------------------ state accessors
  get doc() { return this.app.doc; }
  get wb() { return this.app.doc?.wb; }
  get sheet() { return this.wb?.sheet; }
  get view() { return this.app.view(); }
  get zoom() { return this.view?.zoom || 1; }

  hdrW() {
    const g = geom(this.sheet);
    const lastRow = Math.min(MAX_ROWS, g.rows.at(this.scroller.scrollTop / this.zoom + this.host.clientHeight / this.zoom) + 1);
    const digits = String(lastRow).length;
    return Math.max(36, digits * 8 + 14);
  }

  metrics() {
    const sh = this.sheet, g = geom(sh), z = this.zoom;
    const HW = this.hdrW(), HH = HDR_H;
    const fr = sh.freeze?.r || 0, fc = sh.freeze?.c || 0;
    const FW = g.cols.pos(fc), FH = g.rows.pos(fr);
    const W = this.scroller.clientWidth / z, H = this.scroller.clientHeight / z;
    const sx = this.scroller.scrollLeft / z, sy = this.scroller.scrollTop / z;
    return { sh, g, z, HW, HH, fr, fc, FW, FH, W, H, sx, sy };
  }

  xOf(c, m = this.metrics()) { return m.HW + m.g.cols.pos(c) - (c >= m.fc ? m.sx : 0); }
  yOf(r, m = this.metrics()) { return m.HH + m.g.rows.pos(r) - (r >= m.fr ? m.sy : 0); }
  colAt(x, m = this.metrics()) {
    if (m.fc && x < m.HW + m.FW) return m.g.cols.at(Math.max(0, x - m.HW));
    return m.g.cols.at(Math.max(m.FW, x - m.HW + m.sx));
  }
  rowAt(y, m = this.metrics()) {
    if (m.fr && y < m.HH + m.FH) return m.g.rows.at(Math.max(0, y - m.HH));
    return m.g.rows.at(Math.max(m.FH, y - m.HH + m.sy));
  }

  /** screen rectangle (CSS px relative to host) of a cell or its merge */
  cellRect(r, c, { merged = true } = {}) {
    const m = this.metrics();
    let r2 = r, c2 = c;
    if (merged) { const mg = m.sh.mergeAt(r, c); if (mg) { r = mg.r1; c = mg.c1; r2 = mg.r2; c2 = mg.c2; } }
    const x = this.xOf(c, m), y = this.yOf(r, m);
    const w = m.g.cols.pos(c2 + 1) - m.g.cols.pos(c), h = m.g.rows.pos(r2 + 1) - m.g.rows.pos(r);
    return { x: x * m.z, y: y * m.z, w: w * m.z, h: h * m.z, z: m.z };
  }

  // ------------------------------------------------------------------ display helpers
  fontFor(st) {
    const wb = this.wb;
    const size = ptToPx(st?.fs || wb.defaultFont.size || 11);
    const fam = st?.ff || wb.defaultFont.name || 'Calibri';
    return { css: `${st?.i ? 'italic ' : ''}${st?.b ? 'bold ' : ''}${size}px "${fam}", ${FONT_FALLBACK}`, size };
  }

  display(sh, r, c) {
    const cell = sh.cells.get(key(r, c));
    if (!cell) return null;
    const st = cell.s;
    let v = cell.f ? this.wb.value(sh, r, c) : (cell.t === 'e' ? { err: cell.v } : cell.v);
    if (v === undefined || v === null) {
      if (cell.f && this.view.showFormulas) return { text: cell.f, align: 'left', st, cell, isNum: false };
      return cell.f ? { text: '', st, cell, align: 'left' } : (st || cell.cm ? { text: '', st, cell, align: 'left' } : null);
    }
    if (this.view.showFormulas && cell.f) return { text: cell.f, align: 'left', st, cell, isNum: false };
    const { text, color } = formatValue(v, st?.nf);
    const isNum = typeof v === 'number';
    let align = st?.ha;
    if (!align || align === 'general') align = isNum ? 'right' : (typeof v === 'boolean' || (v && v.err)) ? 'center' : 'left';
    if (align === 'centerContinuous' || align === 'distributed') align = 'center';
    if (align === 'fill' || align === 'justify') align = 'left';
    return { text, color: color ? COLOR_NAMES[color] : null, align, st, cell, isNum, value: v };
  }

  measure(font, text) {
    const k = font + '\u0000' + text;
    let w = this.measureCache.get(k);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.measureCache.size > 20000) this.measureCache.clear();
      this.measureCache.set(k, w);
    }
    return w;
  }

  // ------------------------------------------------------------------ rendering
  render() {
    if (this._raf) return;
    this._raf = requestAnimationFrame(() => { this._raf = 0; this.draw(); });
  }

  draw() {
    if (!this.sheet) return;
    const dpr = window.devicePixelRatio || 1;
    const cw = this.scroller.clientWidth, ch = this.scroller.clientHeight;
    if (this.canvas.width !== Math.round(cw * dpr) || this.canvas.height !== Math.round(ch * dpr)) {
      this.canvas.width = Math.round(cw * dpr); this.canvas.height = Math.round(ch * dpr);
      this.canvas.style.width = cw + 'px'; this.canvas.style.height = ch + 'px';
    }
    const m = this.metrics();
    this.updateExtent(m);
    const ctx = this.ctx;
    ctx.setTransform(dpr * m.z, 0, 0, dpr * m.z, 0, 0);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, m.W, m.H);
    const g = m.g;
    const sc0 = g.cols.at(m.FW + m.sx), sc1 = g.cols.at(m.W - m.HW + m.sx);
    const sr0 = g.rows.at(m.FH + m.sy), sr1 = g.rows.at(m.H - m.HH + m.sy);
    const regions = [
      { r0: sr0, r1: sr1, c0: sc0, c1: sc1, clip: [m.HW + m.FW, m.HH + m.FH, m.W, m.H] },
    ];
    if (m.fr) regions.push({ r0: 0, r1: m.fr - 1, c0: sc0, c1: sc1, clip: [m.HW + m.FW, m.HH, m.W, m.HH + m.FH] });
    if (m.fc) regions.push({ r0: sr0, r1: sr1, c0: 0, c1: m.fc - 1, clip: [m.HW, m.HH + m.FH, m.HW + m.FW, m.H] });
    if (m.fr && m.fc) regions.push({ r0: 0, r1: m.fr - 1, c0: 0, c1: m.fc - 1, clip: [m.HW, m.HH, m.HW + m.FW, m.HH + m.FH] });
    for (const rg of regions) this.drawRegion(ctx, m, rg);
    for (const rg of regions) this.drawOverlay(ctx, m, rg);
    this.drawHeaders(ctx, m, sc0, sc1, sr0, sr1);
    if (m.fr || m.fc) {
      ctx.strokeStyle = '#7f868c'; ctx.lineWidth = 1 / m.z; ctx.beginPath();
      if (m.fr) { ctx.moveTo(0, m.HH + m.FH); ctx.lineTo(m.W, m.HH + m.FH); }
      if (m.fc) { ctx.moveTo(m.HW + m.FW, 0); ctx.lineTo(m.HW + m.FW, m.H); }
      ctx.stroke();
    }
    this.lastVisible = { r0: sr0, r1: sr1, c0: sc0, c1: sc1 };
    this.app.charts?.layout();
    this.app.editor?.reposition();
  }

  updateExtent(m) {
    const v = this.view;
    const visR = m.g.rows.at(m.H - m.HH + m.sy) + 1, visC = m.g.cols.at(m.W - m.HW + m.sx) + 1;
    let nR = Math.max(m.sh.maxRow + 50, visR + 30, v.extR || 0, 100);
    let nC = Math.max(m.sh.maxCol + 10, visC + 8, v.extC || 0, 26);
    for (const chrt of m.sh.charts) { nR = Math.max(nR, chrt.anchor.r + 40); nC = Math.max(nC, chrt.anchor.c + 15); }
    nR = Math.min(nR, MAX_ROWS); nC = Math.min(nC, MAX_COLS);
    v.extR = nR; v.extC = nC;
    const w = (m.HW + m.g.cols.pos(nC)) * m.z, h = (m.HH + m.g.rows.pos(nR)) * m.z;
    // browsers cap element size (~33M px in Chromium); clamp
    const cw = Math.min(w, 3.3e7), ch = Math.min(h, 3.3e7);
    if (this.spacer._w !== cw || this.spacer._h !== ch) {
      this.spacer.style.width = cw + 'px'; this.spacer.style.height = ch + 'px';
      this.spacer._w = cw; this.spacer._h = ch;
      this.chartLayer.style.width = cw + 'px'; this.chartLayer.style.height = ch + 'px';
    }
  }

  drawRegion(ctx, m, rg) {
    const { sh, g } = m;
    const [x0, y0, x1, y1] = rg.clip;
    if (x1 <= x0 || y1 <= y0) return;
    ctx.save();
    ctx.beginPath(); ctx.rect(x0, y0, x1 - x0, y1 - y0); ctx.clip();
    const rows = [], cols = [];
    for (let r = rg.r0; r <= rg.r1 && r < MAX_ROWS; r++) { const h = g.rows.size(r); if (h > 0) rows.push(r); }
    for (let c = rg.c0; c <= rg.c1 && c < MAX_COLS; c++) { const w = g.cols.size(c); if (w > 0) cols.push(c); }
    const X = (c) => this.xOf(c, m), Y = (r) => this.yOf(r, m);

    // merged areas in view
    const merges = sh.merges.filter((mg) => mg.r2 >= rg.r0 && mg.r1 <= rg.r1 && mg.c2 >= rg.c0 && mg.c1 <= rg.c1);
    const covered = new Set();
    for (const mg of merges) for (let r = Math.max(mg.r1, rg.r0); r <= Math.min(mg.r2, rg.r1); r++) for (let c = Math.max(mg.c1, rg.c0); c <= Math.min(mg.c2, rg.c1); c++) covered.add(key(r, c));

    // 1) fills
    for (const r of rows) {
      const y = Y(r), h = g.rows.size(r);
      for (const c of cols) {
        const cell = sh.cells.get(key(r, c));
        const bg = cell?.s?.bg;
        if (bg && !covered.has(key(r, c))) { ctx.fillStyle = bg; ctx.fillRect(X(c), y, g.cols.size(c), h); }
      }
    }
    // 2) gridlines
    if (sh.showGrid !== false) {
      ctx.strokeStyle = GRID_COLOR; ctx.lineWidth = 1 / m.z; ctx.beginPath();
      const half = 0.5 / m.z;
      for (const c of cols) { const x = X(c + 1) - half; ctx.moveTo(x, y0); ctx.lineTo(x, y1); }
      for (const r of rows) { const y = Y(r + 1) - half; ctx.moveTo(x0, y); ctx.lineTo(x1, y); }
      ctx.stroke();
      // re-cover gridlines under filled cells
      for (const r of rows) {
        const y = Y(r), h = g.rows.size(r);
        for (const c of cols) {
          const bg = sh.cells.get(key(r, c))?.s?.bg;
          const nb = sh.cells.get(key(r, c + 1))?.s?.bg;
          if (bg && nb && !covered.has(key(r, c))) { ctx.fillStyle = bg; ctx.fillRect(X(c + 1) - 1, y, 1, h - 1); }
        }
      }
    }
    // merged backgrounds (hide interior gridlines)
    for (const mg of merges) {
      const x = X(mg.c1), y = Y(mg.r1), w = g.cols.pos(mg.c2 + 1) - g.cols.pos(mg.c1), h = g.rows.pos(mg.r2 + 1) - g.rows.pos(mg.r1);
      const st = sh.cells.get(key(mg.r1, mg.c1))?.s;
      ctx.fillStyle = st?.bg || '#fff';
      ctx.fillRect(x, y, w - 1 / m.z, h - 1 / m.z);
    }

    // 3) text
    ctx.textBaseline = 'alphabetic';
    const OVER = 12;
    const cStart = Math.max(0, rg.c0 - OVER), cEnd = Math.min(MAX_COLS - 1, rg.c1 + OVER);
    for (const r of rows) {
      const y = Y(r), h = g.rows.size(r);
      for (let c = cStart; c <= cEnd; c++) {
        if (g.cols.size(c) === 0) continue;
        const k = key(r, c);
        if (covered.has(k)) continue;
        if ((c < rg.c0 || c > rg.c1) && sh.merges.length && sh.mergeAt(r, c)) continue;
        const d = this.display(sh, r, c);
        if (!d) continue;
        if (d.cell.cm) this.drawCommentMark(ctx, X(c) + g.cols.size(c), y);
        if (!d.text) continue;
        this.drawText(ctx, m, d, r, c, X(c), y, g.cols.size(c), h, rows, true);
      }
    }
    for (const mg of merges) {
      const d = this.display(sh, mg.r1, mg.c1);
      if (!d || !d.text) continue;
      const x = X(mg.c1), y = Y(mg.r1), w = g.cols.pos(mg.c2 + 1) - g.cols.pos(mg.c1), h = g.rows.pos(mg.r2 + 1) - g.rows.pos(mg.r1);
      this.drawText(ctx, m, d, mg.r1, mg.c1, x, y, w, h, rows, false);
    }

    // 4) borders
    for (const r of rows) {
      const y = Y(r), h = g.rows.size(r);
      for (const c of cols) {
        const st = sh.cells.get(key(r, c))?.s;
        if (!st || !(st.bl || st.br || st.bt || st.bb)) continue;
        const mg = sh.mergeAt(r, c);
        const x = X(c), w = g.cols.size(c);
        if (st.bt && (!mg || r === mg.r1)) this.line(ctx, m, st.bt, x, y, x + w, y, 'h');
        if (st.bb && (!mg || r === mg.r2)) this.line(ctx, m, st.bb, x, y + h, x + w, y + h, 'h');
        if (st.bl && (!mg || c === mg.c1)) this.line(ctx, m, st.bl, x, y, x, y + h, 'v');
        if (st.br && (!mg || c === mg.c2)) this.line(ctx, m, st.br, x + w, y, x + w, y + h, 'v');
      }
    }
    // merged top-left border spanning
    for (const mg of merges) {
      const st = sh.cells.get(key(mg.r1, mg.c1))?.s;
      if (!st) continue;
      const x = X(mg.c1), y = Y(mg.r1), w = g.cols.pos(mg.c2 + 1) - g.cols.pos(mg.c1), h = g.rows.pos(mg.r2 + 1) - g.rows.pos(mg.r1);
      if (st.bt) this.line(ctx, m, st.bt, x, y, x + w, y, 'h');
      if (st.bl) this.line(ctx, m, st.bl, x, y, x, y + h, 'v');
    }
    // 5) autofilter buttons
    const af = sh.autoFilter;
    if (af && rows.includes(af.range.r1)) {
      for (let c = Math.max(af.range.c1, rg.c0); c <= Math.min(af.range.c2, rg.c1); c++) {
        if (g.cols.size(c) === 0) continue;
        const active = !!af.cols[String(c - af.range.c1)];
        this.drawFilterButton(ctx, X(c) + g.cols.size(c), Y(af.range.r1), g.rows.size(af.range.r1), active);
      }
    }
    ctx.restore();
  }

  drawCommentMark(ctx, xr, y) {
    ctx.fillStyle = '#d93025';
    ctx.beginPath(); ctx.moveTo(xr - 6, y); ctx.lineTo(xr - 1, y); ctx.lineTo(xr - 1, y + 5); ctx.closePath(); ctx.fill();
  }

  drawFilterButton(ctx, xr, y, h, active) {
    const s = Math.min(16, h - 3), bx = xr - s - 2, by = y + (h - s) / 2;
    ctx.fillStyle = active ? '#e6f4ea' : '#fff';
    ctx.strokeStyle = active ? ACCENT : '#a0a4a8';
    ctx.lineWidth = 1;
    ctx.fillRect(bx, by, s, s); ctx.strokeRect(bx + 0.5, by + 0.5, s - 1, s - 1);
    ctx.fillStyle = active ? ACCENT : '#555';
    ctx.beginPath();
    const cx = bx + s / 2, cy = by + s / 2;
    if (active) { // funnel
      ctx.moveTo(cx - 5, cy - 4); ctx.lineTo(cx + 5, cy - 4); ctx.lineTo(cx + 1, cy + 1); ctx.lineTo(cx + 1, cy + 5); ctx.lineTo(cx - 1, cy + 4); ctx.lineTo(cx - 1, cy + 1); ctx.closePath();
    } else { ctx.moveTo(cx - 4, cy - 2); ctx.lineTo(cx + 4, cy - 2); ctx.lineTo(cx, cy + 3); ctx.closePath(); }
    ctx.fill();
  }

  line(ctx, m, b, xa, ya, xb, yb, dir) {
    const w = BORDER_W[b.s] || 1;
    ctx.save();
    ctx.strokeStyle = b.c || '#000';
    ctx.lineWidth = (b.s === 'double' ? 1 : w) / (w === 1 ? m.z : 1);
    ctx.setLineDash(BORDER_DASH[b.s] || []);
    const off = (w % 2 === 1 ? 0.5 : 0) / m.z;
    ctx.beginPath();
    if (b.s === 'double') {
      const d = 1.5 / m.z;
      if (dir === 'h') { ctx.moveTo(xa, ya - d + off); ctx.lineTo(xb, yb - d + off); ctx.moveTo(xa, ya + d - off); ctx.lineTo(xb, yb + d - off); }
      else { ctx.moveTo(xa - d + off, ya); ctx.lineTo(xb - d + off, yb); ctx.moveTo(xa + d - off, ya); ctx.lineTo(xb + d - off, yb); }
    } else if (dir === 'h') { ctx.moveTo(xa - (w > 1 ? w / 2 : 0), ya - off); ctx.lineTo(xb + (w > 1 ? w / 2 : 0), yb - off); }
    else { ctx.moveTo(xa - off, ya); ctx.lineTo(xb - off, yb); }
    ctx.stroke();
    ctx.restore();
  }

  drawText(ctx, m, d, r, c, x, y, w, h, rows, allowOverflow) {
    const { sh, g } = m;
    const st = d.st || {};
    const font = this.fontFor(st);
    ctx.font = font.css;
    let text = d.text;
    const pad = 3;
    const indent = (st.ind || 0) * 9;
    const wrap = st.wrap || text.includes('\n');
    ctx.fillStyle = d.color || st.fc || '#000';
    let lines = [text];
    let left = x, right = x + w;
    if (wrap) {
      lines = this.wrapLines(font.css, text, Math.max(4, w - pad * 2 - indent));
    } else {
      const tw = this.measure(font.css, text);
      if (tw > w - pad * 2 - indent) {
        if (d.isNum && !this.view.showFormulas) {
          // numbers never overflow: shrink General precision or show ###
          if (!st.nf || st.nf === 'General') {
            let s = text;
            const v = d.value;
            for (let p = 10; p >= 1 && this.measure(font.css, s) > w - pad * 2; p--) s = String(parseFloat(v.toPrecision(p)));
            text = this.measure(font.css, s) <= w - pad * 2 ? s : '#'.repeat(Math.max(1, Math.floor((w - pad * 2) / this.measure(font.css, '#'))));
          } else text = '#'.repeat(Math.max(1, Math.floor((w - pad * 2) / this.measure(font.css, '#'))));
          lines = [text];
        } else if (allowOverflow) {
          // overflow into empty neighbours
          const need = tw + pad * 2 + indent;
          if (d.align === 'left' || d.align === 'center') {
            let cc = c + 1;
            while (right - x < need && cc < MAX_COLS && !this.occupied(sh, r, cc)) { right += g.cols.size(cc); cc++; }
          }
          if (d.align === 'right' || d.align === 'center') {
            let cc = c - 1;
            while (right - left < need && cc >= 0 && !this.occupied(sh, r, cc)) { left -= g.cols.size(cc); cc--; }
          }
        }
      }
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(left + 1 / m.z, y, right - left - 2 / m.z, h - 1 / m.z); ctx.clip();
    const lh = font.size * 1.25;
    const total = lines.length * lh;
    const va = st.va || 'bottom';
    let ty = va === 'top' ? y + 2 : va === 'center' ? y + (h - total) / 2 : y + h - total - 3;
    ty = Math.max(ty, y + (wrap ? 1 : -4));
    for (const ln of lines) {
      const tw = this.measure(font.css, ln);
      let tx;
      if (d.align === 'right') tx = (wrap ? x + w : right) - pad - tw - indent;
      else if (d.align === 'center') tx = (wrap ? x : left) + ((wrap ? w : right - left) - tw) / 2;
      else tx = (wrap ? x : left) + pad + indent;
      const base = ty + lh - (lh - font.size) / 2 - font.size * 0.12;
      ctx.fillText(ln, tx, base);
      if (st.u || (d.cell.h && !st.fc && false)) { ctx.fillRect(tx, base + 1.5, tw, st.u === 'double' ? 0.8 : 1); if (st.u === 'double') ctx.fillRect(tx, base + 3.5, tw, 0.8); }
      if (st.st) ctx.fillRect(tx, base - font.size * 0.3, tw, 1);
      ty += lh;
    }
    ctx.restore();
  }

  occupied(sh, r, c) {
    const cell = sh.cells.get(key(r, c));
    if (sh.mergeAt(r, c)) return true;
    if (!cell) return false;
    if (cell.f) { const v = this.wb.value(sh, r, c); return v !== null && v !== '' && v !== undefined; }
    return cell.v !== undefined && cell.v !== null && cell.v !== '';
  }

  wrapLines(font, text, maxW) {
    const out = [];
    for (const para of String(text).split('\n')) {
      if (!para) { out.push(''); continue; }
      // break on spaces; fall back to character breaking (CJK)
      const tokens = para.match(/\S+\s*|\s+/g) || [para];
      let line = '';
      for (const tk of tokens) {
        if (this.measure(font, line + tk) <= maxW) { line += tk; continue; }
        if (line) { out.push(line.trimEnd()); line = ''; }
        if (this.measure(font, tk) <= maxW) { line = tk; continue; }
        for (const ch of tk) {
          if (this.measure(font, line + ch) > maxW && line) { out.push(line); line = ''; }
          line += ch;
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  drawOverlay(ctx, m, rg) {
    const v = this.view, { g } = m;
    const [x0, y0, x1, y1] = rg.clip;
    if (x1 <= x0 || y1 <= y0) return;
    ctx.save();
    ctx.beginPath(); ctx.rect(x0, y0, x1 - x0, y1 - y0); ctx.clip();
    const rectOf = (rgn) => {
      const x = this.xOf(rgn.c1, m), y = this.yOf(rgn.r1, m);
      return [x, y, g.cols.pos(rgn.c2 + 1) - g.cols.pos(rgn.c1), g.rows.pos(rgn.r2 + 1) - g.rows.pos(rgn.r1)];
    };
    // formula reference highlights
    for (const h of this.refHighlights) {
      if (h.sheet !== m.sh) continue;
      const [x, y, w, hh] = rectOf(h.range);
      ctx.fillStyle = h.color + '1f'; ctx.fillRect(x, y, w, hh);
      ctx.strokeStyle = h.color; ctx.lineWidth = 2 / m.z; ctx.setLineDash([]);
      ctx.strokeRect(x + 1 / m.z, y + 1 / m.z, w - 2 / m.z, hh - 2 / m.z);
    }
    // cells changed by another client (agent / other window) flash briefly
    const now = performance.now();
    for (const f of this.flashes || []) {
      if (f.sheet !== m.sh || f.until < now) continue;
      const [fx, fy, fw, fh] = rectOf(f);
      const a = Math.min(1, (f.until - now) / 1600);
      ctx.fillStyle = `rgba(255,193,7,${0.35 * a})`; ctx.fillRect(fx, fy, fw, fh);
      ctx.strokeStyle = `rgba(230,145,0,${a})`; ctx.lineWidth = 2 / m.z; ctx.setLineDash([]);
      ctx.strokeRect(fx + 1 / m.z, fy + 1 / m.z, fw - 2 / m.z, fh - 2 / m.z);
    }
    const sel = v.sel;
    const [sx, sy, sw, sh] = rectOf(sel);
    const single = sel.r1 === sel.r2 && sel.c1 === sel.c2 || (m.sh.mergeAt(sel.r1, sel.c1) && (() => { const mg = m.sh.mergeAt(sel.r1, sel.c1); return mg.r1 === sel.r1 && mg.c1 === sel.c1 && mg.r2 === sel.r2 && mg.c2 === sel.c2; })());
    if (!single) {
      ctx.fillStyle = 'rgba(33,115,70,0.10)';
      ctx.fillRect(sx, sy, sw, sh);
      const ar = this.cellRect(v.active.r, v.active.c);
      // active cell: repaint without tint
      const ax = ar.x / m.z, ay = ar.y / m.z, aw = ar.w / m.z, ah = ar.h / m.z;
      ctx.save(); ctx.beginPath(); ctx.rect(ax, ay, aw, ah); ctx.clip();
      ctx.fillStyle = '#fff'; ctx.fillRect(ax, ay, aw, ah);
      this.drawRegion(ctx, m, { r0: v.active.r, r1: v.active.r, c0: v.active.c, c1: v.active.c, clip: [Math.max(ax, x0), Math.max(ay, y0), Math.min(ax + aw, x1), Math.min(ay + ah, y1)] });
      ctx.restore();
      ctx.strokeStyle = ACCENT; ctx.lineWidth = 1 / m.z;
      ctx.strokeRect(ax + 1 / m.z, ay + 1 / m.z, aw - 2 / m.z, ah - 2 / m.z);
    }
    ctx.strokeStyle = ACCENT; ctx.lineWidth = 2 / m.z; ctx.setLineDash([]);
    ctx.strokeRect(sx, sy, sw - 1 / m.z, sh - 1 / m.z);
    // fill handle
    if (!this.app.editor?.active) {
      ctx.fillStyle = ACCENT; ctx.strokeStyle = '#fff'; ctx.lineWidth = 1 / m.z;
      const hx = sx + sw - 4 / m.z, hy = sy + sh - 4 / m.z;
      ctx.fillRect(hx, hy, 6 / m.z, 6 / m.z); ctx.strokeRect(hx, hy, 6 / m.z, 6 / m.z);
    }
    // fill preview
    if (this.fillPreview) {
      const [fx, fy, fw, fh] = rectOf(this.fillPreview);
      ctx.strokeStyle = '#555'; ctx.setLineDash([3 / m.z, 2 / m.z]); ctx.lineWidth = 1 / m.z;
      ctx.strokeRect(fx + 0.5, fy + 0.5, fw - 1, fh - 1);
    }
    // copy marquee
    const clip = this.app.clipboard?.state;
    if (clip && clip.sheet === m.sh && clip.wb === this.wb) {
      const [cx, cy, cw2, ch2] = rectOf(clip.range);
      ctx.strokeStyle = ACCENT; ctx.lineWidth = 2 / m.z; ctx.setLineDash([4 / m.z, 3 / m.z]);
      ctx.lineDashOffset = this.marchOffset;
      ctx.strokeRect(cx + 1 / m.z, cy + 1 / m.z, cw2 - 2 / m.z, ch2 - 2 / m.z);
    }
    ctx.restore();
  }

  drawHeaders(ctx, m, sc0, sc1, sr0, sr1) {
    const { g, sh } = m;
    const sel = this.view.sel;
    ctx.font = `11px ${FONT_FALLBACK}`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    const fullCols = sel.r1 === 0 && sel.r2 === MAX_ROWS - 1;
    const fullRows = sel.c1 === 0 && sel.c2 === MAX_COLS - 1;
    // column header
    const colRanges = [[sc0, sc1, m.HW + m.FW, m.W]];
    if (m.fc) colRanges.push([0, m.fc - 1, m.HW, m.HW + m.FW]);
    for (const [c0, c1, cx0, cx1] of colRanges) {
      ctx.save(); ctx.beginPath(); ctx.rect(cx0, 0, cx1 - cx0, m.HH); ctx.clip();
      ctx.fillStyle = HDR_BG; ctx.fillRect(cx0, 0, cx1 - cx0, m.HH);
      for (let c = c0; c <= c1; c++) {
        const w = g.cols.size(c); if (!w) continue;
        const x = this.xOf(c, m);
        const on = c >= sel.c1 && c <= sel.c2;
        if (on) { ctx.fillStyle = fullCols ? '#c4e3d1' : HDR_SEL_BG; ctx.fillRect(x, 0, w, m.HH); ctx.fillStyle = ACCENT; ctx.fillRect(x, m.HH - 2, w, 2); }
        ctx.fillStyle = on ? ACCENT : '#444';
        ctx.fillText(colName(c), x + w / 2, m.HH / 2 + 1);
        ctx.fillStyle = '#d0d4d8'; ctx.fillRect(x + w - 1 / m.z, 0, 1 / m.z, m.HH);
      }
      ctx.restore();
    }
    // row header
    ctx.textAlign = 'right';
    const rowRanges = [[sr0, sr1, m.HH + m.FH, m.H]];
    if (m.fr) rowRanges.push([0, m.fr - 1, m.HH, m.HH + m.FH]);
    for (const [r0, r1, ry0, ry1] of rowRanges) {
      ctx.save(); ctx.beginPath(); ctx.rect(0, ry0, m.HW, ry1 - ry0); ctx.clip();
      ctx.fillStyle = HDR_BG; ctx.fillRect(0, ry0, m.HW, ry1 - ry0);
      for (let r = r0; r <= r1; r++) {
        const h = g.rows.size(r); if (!h) continue;
        const y = this.yOf(r, m);
        const on = r >= sel.r1 && r <= sel.r2;
        if (on) { ctx.fillStyle = fullRows ? '#c4e3d1' : HDR_SEL_BG; ctx.fillRect(0, y, m.HW, h); ctx.fillStyle = ACCENT; ctx.fillRect(m.HW - 2, y, 2, h); }
        ctx.fillStyle = sh.filterHidden.size && sh.autoFilter && r > sh.autoFilter.range.r1 && r <= sh.autoFilter.range.r2 ? '#1a73e8' : on ? ACCENT : '#444';
        if (h > 8) ctx.fillText(String(r + 1), m.HW - 6, y + h / 2 + 1);
        ctx.fillStyle = '#d0d4d8'; ctx.fillRect(0, y + h - 1 / m.z, m.HW, 1 / m.z);
        // hidden-row marker
        if (g.rows.size(r + 1) === 0 && r + 1 < MAX_ROWS) { ctx.fillStyle = '#888'; ctx.fillRect(0, y + h - 1.5, m.HW, 1.5); }
      }
      ctx.restore();
    }
    // corner
    ctx.fillStyle = HDR_BG; ctx.fillRect(0, 0, m.HW, m.HH);
    ctx.fillStyle = '#b8bcc0';
    ctx.beginPath(); ctx.moveTo(m.HW - 4, m.HH - 4); ctx.lineTo(m.HW - 4, 6); ctx.lineTo(m.HW - 14, m.HH - 4); ctx.closePath(); ctx.fill();
    ctx.fillStyle = '#c8ccd0';
    ctx.fillRect(0, m.HH - 1, m.W, 1); ctx.fillRect(m.HW - 1, 0, 1, m.H);
    ctx.textAlign = 'left';
  }

  // ------------------------------------------------------------------ hit testing
  hit(clientX, clientY) {
    const rect = this.scroller.getBoundingClientRect();
    const m = this.metrics();
    const x = (clientX - rect.left) / m.z, y = (clientY - rect.top) / m.z;
    if (x > m.W || y > m.H) return { type: 'scrollbar' };
    const res = { x, y, m };
    if (x < m.HW && y < m.HH) return { ...res, type: 'corner' };
    if (y < m.HH) {
      const c = this.colAt(x, m);
      const xl = this.xOf(c, m), xr = xl + m.g.cols.size(c);
      if (xr - x < 4) return { ...res, type: 'colresize', c };
      if (x - xl < 4 && c > 0) { let p = c - 1; return { ...res, type: 'colresize', c: p }; }
      return { ...res, type: 'colhdr', c };
    }
    if (x < m.HW) {
      const r = this.rowAt(y, m);
      const yt = this.yOf(r, m), yb = yt + m.g.rows.size(r);
      if (yb - y < 3) return { ...res, type: 'rowresize', r };
      if (y - yt < 3 && r > 0) return { ...res, type: 'rowresize', r: r - 1 };
      return { ...res, type: 'rowhdr', r };
    }
    const r = this.rowAt(y, m), c = this.colAt(x, m);
    const v = this.view;
    // fill handle
    const sx2 = (this.xOf(v.sel.c2 + 1, m)) , sy2 = (this.yOf(v.sel.r2 + 1, m));
    if (Math.abs(x - sx2) < 5 && Math.abs(y - sy2) < 5) return { ...res, type: 'fill', r, c };
    // filter button
    const af = m.sh.autoFilter;
    if (af && r === af.range.r1 && c >= af.range.c1 && c <= af.range.c2) {
      const xr = this.xOf(c, m) + m.g.cols.size(c);
      if (xr - x < 19) return { ...res, type: 'filter', r, c };
    }
    return { ...res, type: 'cell', r, c };
  }

  // ------------------------------------------------------------------ interactions
  bind() {
    const sc = this.scroller;
    sc.addEventListener('scroll', () => { this.render(); this.app.onScroll?.(); });
    sc.addEventListener('wheel', (e) => {
      if (e.ctrlKey || e.metaKey) { e.preventDefault(); this.app.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 0.9)); }
    }, { passive: false });
    sc.addEventListener('pointerdown', (e) => this.onDown(e));
    sc.addEventListener('pointermove', (e) => this.onHover(e));
    sc.addEventListener('dblclick', (e) => this.onDbl(e));
    sc.addEventListener('contextmenu', (e) => {
      if (e.target.closest('.chart-box')) return;
      e.preventDefault();
      const h = this.hit(e.clientX, e.clientY);
      if (h.type === 'cell' && !inRange(this.view.sel, h.r, h.c)) this.app.select(h.r, h.c);
      if (h.type === 'colhdr' && !(this.view.sel.c1 <= h.c && h.c <= this.view.sel.c2 && this.view.sel.r1 === 0 && this.view.sel.r2 === MAX_ROWS - 1)) this.app.selectRange({ r1: 0, c1: h.c, r2: MAX_ROWS - 1, c2: h.c });
      if (h.type === 'rowhdr' && !(this.view.sel.r1 <= h.r && h.r <= this.view.sel.r2 && this.view.sel.c1 === 0 && this.view.sel.c2 === MAX_COLS - 1)) this.app.selectRange({ r1: h.r, c1: 0, r2: h.r, c2: MAX_COLS - 1 });
      this.app.contextMenu(e.clientX, e.clientY, h.type);
    });
  }

  onHover(e) {
    if (this.drag) return;
    if (e.target.closest('.chart-box')) return;
    const h = this.hit(e.clientX, e.clientY);
    const cur = { colresize: 'col-resize', rowresize: 'row-resize', fill: 'crosshair', colhdr: 's-resize', rowhdr: 'e-resize', filter: 'pointer', corner: 'default' }[h.type] || 'cell';
    this.scroller.style.cursor = cur;
    const cell = h.type === 'cell' ? this.sheet.get(h.r, h.c) : null;
    this.app.hoverComment?.(cell?.cm ? { text: cell.cm, x: e.clientX, y: e.clientY } : null);
  }

  onDbl(e) {
    if (e.target.closest('.chart-box')) return;
    const h = this.hit(e.clientX, e.clientY);
    if (h.type === 'colresize') return this.app.autofitCols(h.c);
    if (h.type === 'rowresize') return this.app.autofitRows(h.r);
    if (h.type === 'cell') this.app.editor.start({ mode: 'edit' });
  }

  onDown(e) {
    if (e.target.closest('.chart-box')) return;
    if (e.button === 2) return;
    const h = this.hit(e.clientX, e.clientY);
    if (h.type === 'scrollbar') return;
    this.app.charts?.deselect();
    const ed = this.app.editor;
    // pointing mode: clicking cells while editing a formula inserts references
    if (ed.active && ed.isFormulaPointable() && (h.type === 'cell' || h.type === 'colhdr' || h.type === 'rowhdr')) {
      e.preventDefault();
      const start = h.type === 'cell' ? { r: h.r, c: h.c } : h.type === 'colhdr' ? { r: 0, c: h.c } : { r: h.r, c: 0 };
      const whole = h.type === 'colhdr' ? 'col' : h.type === 'rowhdr' ? 'row' : null;
      ed.pointRange({ r1: start.r, c1: start.c, r2: whole === 'col' ? MAX_ROWS - 1 : start.r, c2: whole === 'row' ? MAX_COLS - 1 : start.c }, true, this.sheet);
      this.dragLoop(e, (hh) => {
        if (hh.r === undefined && hh.c === undefined) return;
        const rr = hh.r ?? start.r, cc = hh.c ?? start.c;
        ed.pointRange(normRange({ r1: whole === 'col' ? 0 : start.r, c1: whole === 'row' ? 0 : start.c, r2: whole === 'col' ? MAX_ROWS - 1 : rr, c2: whole === 'row' ? MAX_COLS - 1 : cc }), false, this.sheet);
      });
      return;
    }
    if (ed.active) { if (!ed.commit()) return; }
    this.app.focusGrid();
    const v = this.view;
    switch (h.type) {
      case 'corner': this.app.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, { r: 0, c: 0 }); break;
      case 'colhdr': {
        const a = e.shiftKey ? v.anchor.c : h.c;
        this.app.selectRange({ r1: 0, c1: Math.min(a, h.c), r2: MAX_ROWS - 1, c2: Math.max(a, h.c) }, { r: v.sel.r1 === 0 && e.shiftKey ? v.active.r : 0, c: a }, { anchor: { r: 0, c: a } });
        this.dragLoop(e, (hh) => { if (hh.c === undefined) return; this.app.selectRange({ r1: 0, c1: Math.min(a, hh.c), r2: MAX_ROWS - 1, c2: Math.max(a, hh.c) }, { r: 0, c: a }, { anchor: { r: 0, c: a }, scroll: false }); }, 'x');
        break;
      }
      case 'rowhdr': {
        const a = e.shiftKey ? v.anchor.r : h.r;
        this.app.selectRange({ r1: Math.min(a, h.r), c1: 0, r2: Math.max(a, h.r), c2: MAX_COLS - 1 }, { r: a, c: 0 }, { anchor: { r: a, c: 0 } });
        this.dragLoop(e, (hh) => { if (hh.r === undefined) return; this.app.selectRange({ r1: Math.min(a, hh.r), c1: 0, r2: Math.max(a, hh.r), c2: MAX_COLS - 1 }, { r: a, c: 0 }, { anchor: { r: a, c: 0 }, scroll: false }); }, 'y');
        break;
      }
      case 'colresize': this.resizeDrag(e, 'col', h.c); break;
      case 'rowresize': this.resizeDrag(e, 'row', h.r); break;
      case 'fill': this.fillDrag(e); break;
      case 'filter': this.app.filters.openMenu(h.c, e.clientX, e.clientY); break;
      case 'cell': {
        if (e.ctrlKey || e.metaKey) {
          const cell = this.sheet.get(h.r, h.c);
          if (cell?.h) { window.open(cell.h, '_blank', 'noopener'); return; }
        }
        if (e.shiftKey) this.app.extendTo(h.r, h.c);
        else this.app.select(h.r, h.c);
        const anchor = { ...this.view.anchor };
        this.dragLoop(e, (hh) => { if (hh.r === undefined) return; this.app.selectRange(normRange({ r1: anchor.r, c1: anchor.c, r2: hh.r, c2: hh.c }), this.view.active, { anchor, scroll: false, expandMerges: true }); });
        break;
      }
      default: break;
    }
  }

  /** generic drag with edge auto-scroll; cb receives {r,c} of pointer cell */
  dragLoop(e, cb, axis) {
    const sc = this.scroller;
    sc.setPointerCapture(e.pointerId);
    let last = e, timer = 0;
    const tick = () => {
      const rect = sc.getBoundingClientRect();
      const m = this.metrics();
      let dx = 0, dy = 0;
      if (last.clientX > rect.left + sc.clientWidth - 4) dx = 20; else if (last.clientX < rect.left + (m.HW + m.FW) * m.z && axis !== 'y') dx = (last.clientX < rect.left + m.HW * m.z ? -20 : 0);
      if (last.clientY > rect.top + sc.clientHeight - 4) dy = 20; else if (last.clientY < rect.top + (m.HH + m.FH) * m.z && axis !== 'x') dy = (last.clientY < rect.top + m.HH * m.z ? -20 : 0);
      if (axis === 'x') dy = 0; if (axis === 'y') dx = 0;
      if (dx || dy) { sc.scrollLeft += dx; sc.scrollTop += dy; move(last); }
    };
    const move = (ev) => {
      last = ev;
      const rect = sc.getBoundingClientRect();
      const m = this.metrics();
      const x = Math.min(Math.max((ev.clientX - rect.left) / m.z, m.HW + 1), m.W - 1);
      const y = Math.min(Math.max((ev.clientY - rect.top) / m.z, m.HH + 1), m.H - 1);
      cb({ r: this.rowAt(y, m), c: this.colAt(x, m) });
    };
    const up = () => {
      clearInterval(timer);
      sc.removeEventListener('pointermove', move);
      sc.removeEventListener('pointerup', up);
      sc.removeEventListener('pointercancel', up);
      this.drag = null;
      this.onDragEnd?.();
      this.onDragEnd = null;
    };
    this.drag = true;
    timer = setInterval(tick, 50);
    sc.addEventListener('pointermove', move);
    sc.addEventListener('pointerup', up);
    sc.addEventListener('pointercancel', up);
  }

  resizeDrag(e, axis, idx) {
    const sc = this.scroller;
    sc.setPointerCapture(e.pointerId);
    const m = this.metrics();
    const start = axis === 'col' ? e.clientX : e.clientY;
    const orig = axis === 'col' ? m.g.cols.size(idx) : m.g.rows.size(idx);
    const sel = this.view.sel;
    const targets = [];
    if (axis === 'col' && sel.r1 === 0 && sel.r2 === MAX_ROWS - 1 && idx >= sel.c1 && idx <= sel.c2) for (let c = sel.c1; c <= sel.c2; c++) targets.push(c);
    else if (axis === 'row' && sel.c1 === 0 && sel.c2 === MAX_COLS - 1 && idx >= sel.r1 && idx <= sel.r2) for (let r = sel.r1; r <= sel.r2; r++) targets.push(r);
    else targets.push(idx);
    const tip = this.app.tooltip;
    let size = orig;
    const move = (ev) => {
      const d = ((axis === 'col' ? ev.clientX : ev.clientY) - start) / m.z;
      size = Math.max(0, Math.round(orig + d));
      // live preview via temporary sizes (not in undo stack)
      const sh = this.sheet;
      for (const t of targets) (axis === 'col' ? sh.colW : sh.rowH).set(t, size);
      sh._geomVersion++;
      tip?.(axis === 'col' ? `너비: ${((size - 5) / 7).toFixed(2)} (${size} 픽셀)` : `높이: ${(size * 0.75).toFixed(2)} (${size} 픽셀)`, ev.clientX + 12, ev.clientY + 12);
      this.render();
    };
    const sh0 = this.sheet;
    const before = sh0.metaSnapshot();
    const up = () => {
      sc.removeEventListener('pointermove', move); sc.removeEventListener('pointerup', up);
      tip?.(null);
      sh0.applyMeta(before);
      this.app.resize(axis, targets, size);
    };
    sc.addEventListener('pointermove', move);
    sc.addEventListener('pointerup', up);
  }

  fillDrag(e) {
    const src = { ...this.view.sel };
    let target = null;
    this.dragLoop(e, (hh) => {
      if (hh.r === undefined) return;
      const { r, c } = hh;
      let t;
      const dr = r > src.r2 ? r - src.r2 : r < src.r1 ? r - src.r1 : 0;
      const dc = c > src.c2 ? c - src.c2 : c < src.c1 ? c - src.c1 : 0;
      if (Math.abs(dr) >= Math.abs(dc) && dr) t = dr > 0 ? { ...src, r2: r } : { ...src, r1: r };
      else if (dc) t = dc > 0 ? { ...src, c2: c } : { ...src, c1: c };
      else t = null;
      target = t;
      this.fillPreview = t;
      this.render();
    });
    this.onDragEnd = () => {
      this.fillPreview = null;
      if (target) this.app.fill(src, target);
      this.render();
    };
  }

  /** scroll so that (r,c) is visible */
  scrollTo(r, c) {
    const m = this.metrics();
    const sc = this.scroller;
    if (c >= m.fc) {
      const x = m.g.cols.pos(c), w = m.g.cols.size(c);
      const viewL = m.FW + m.sx, viewR = m.sx + m.W - m.HW;
      if (x < viewL) sc.scrollLeft = (x - m.FW) * m.z;
      else if (x + w > viewR) sc.scrollLeft = Math.max(0, (x + w - (m.W - m.HW)) * m.z + 2);
    }
    if (r >= m.fr) {
      const y = m.g.rows.pos(r), h = m.g.rows.size(r);
      const viewT = m.FH + m.sy, viewB = m.sy + m.H - m.HH;
      if (y < viewT) sc.scrollTop = (y - m.FH) * m.z;
      else if (y + h > viewB) {
        if (this.view.extR < r + 30) { this.view.extR = r + 60; this.updateExtent(this.metrics()); }
        sc.scrollTop = Math.max(0, (y + h - (m.H - m.HH)) * m.z + 2);
      }
    }
    if (c >= m.fc && this.view.extC < c + 5) { this.view.extC = c + 20; this.updateExtent(this.metrics()); this.scrollTo(r, c); }
    this.render();
  }

  visibleRowCount() {
    const m = this.metrics();
    return Math.max(1, Math.floor((m.H - m.HH - m.FH) / m.g.rows.def) - 1);
  }
}
