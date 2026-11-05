// Chart layer: renders Excel chart models (same series refs as the xlsx) with Chart.js.
import { parseSheetRef, absRangeA1, rangeA1, MAX_ROWS, MAX_COLS } from './address.js';
import { quoteSheetName } from './formula.js';
import { formatValue } from './numfmt.js';
import { newId } from './model.js';
import { geom } from './geometry.js';

// Office 2013-2022 theme accents: the palette when a workbook brings no theme of its own
export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

// Same numbers as server/chart_style.py (the server also sends them with each workbook).
const DEFAULT_STYLE = { text: '#595959', line: '#D9D9D9', titlePt: 14, textPt: 9, seriesLinePt: 2.25, gap: 219, overlap: -27, stackedGap: 150, scatterMarker: 5, pieBorderPt: 1 };
const PX = 4 / 3; // points -> CSS pixels

/** DrawingML lumMod/lumOff on a hex color (HSL luminance), as colors.lum() on the server. */
function lum(hex, mod, off) {
  const n = parseInt(hex.replace('#', ''), 16);
  let r = (n >> 16) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0, l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = Math.max(0, Math.min(1, l * mod / 100000 + off / 100000));
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return '#' + [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Series colors in Excel's order for a workbook theme: accent1-6, darker 60 %, lighter 80 %+20 %. */
export function themePalette(themeColors) {
  const acc = themeColors?.length >= 10 ? themeColors.slice(4, 10).map((c) => '#' + c.replace('#', '')) : PALETTE.slice(0, 6);
  return [...acc, ...acc.map((c) => lum(c, 60000, 0)), ...acc.map((c) => lum(c, 80000, 20000))];
}
export const CHART_TYPES = [
  ['bar', '세로 막대형', { barDir: 'col' }], ['barH', '가로 막대형', { type: 'bar', barDir: 'bar' }],
  ['barStacked', '누적 세로 막대형', { type: 'bar', barDir: 'col', grouping: 'stacked' }],
  ['line', '꺾은선형'], ['area', '영역형'], ['pie', '원형'], ['doughnut', '도넛형'], ['scatter', '분산형'], ['radar', '방사형'],
];

export class ChartLayer {
  constructor(app, grid) {
    this.app = app;
    this.grid = grid;
    this.layer = grid.chartLayer;
    this.instances = new Map(); // chart.id -> {el, chart, sig}
    this.selected = null;
    this.sheetShown = null;
  }

  // ------------------------------------------------------------ data
  resolveRef(ref, defaultSheet) {
    const p = parseSheetRef(ref);
    if (!p) return null;
    const wb = this.app.doc.wb;
    const sh = p.sheet ? wb.sheetByName(p.sheet) : defaultSheet;
    if (!sh) return null;
    const g = p.range;
    const r2 = Math.min(g.r2, Math.max(sh.maxRow, g.r1)), c2 = Math.min(g.c2, Math.max(sh.maxCol, g.c1));
    const vals = [];
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
      if (sh.isRowHidden(r) || sh.isColHidden(c)) continue; // Excel default: hidden cells not plotted
      vals.push({ v: wb.value(sh, r, c), nf: sh.get(r, c)?.s?.nf });
    }
    return vals;
  }

  dataOf(model, sheet) {
    const series = model.series.map((s, i) => {
      const vals = (s.val && this.resolveRef(s.val, sheet)) || [];
      let name = s.name;
      if (s.nameRef) { const nv = this.resolveRef(s.nameRef, sheet); if (nv?.length) name = nv.map((x) => formatValue(x.v, x.nf).text).join(' '); }
      const cats = s.cat ? this.resolveRef(s.cat, sheet) : null;
      return {
        name: name || `계열${i + 1}`,
        type: s.type || model.type,
        values: vals.map((x) => (typeof x.v === 'number' ? x.v : null)),
        cats: cats ? cats.map((x) => (model.type === 'scatter' ? (typeof x.v === 'number' ? x.v : null) : formatValue(x.v, x.nf).text)) : null,
        color: s.color,
        pointColors: s.pointColors,
      };
    });
    const labels = series.find((s) => s.cats)?.cats || (series[0] ? series[0].values.map((_, i) => String(i + 1)) : []);
    return { series, labels };
  }

  config(model, sheet) {
    const { series, labels } = this.dataOf(model, sheet);
    const wb = this.app?.doc?.wb;
    const ST = { ...DEFAULT_STYLE, ...(wb?.chartStyle || {}) };
    const pal = themePalette(wb?.themeColors);
    const t = model.type;
    const pieLike = t === 'pie' || t === 'doughnut';
    const horizontal = t === 'bar' && model.barDir === 'bar';
    const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const pct = model.grouping === 'percentStacked';
    const lineW = ST.seriesLinePt * PX;
    let datasets;
    if (pieLike) {
      const s = series[0] || { values: [], name: '' };
      // a pie with varyColors=0 in the file is one color; otherwise one per slice
      const one = model.varyColors === false ? (s.color || pal[0]) : null;
      datasets = [{ label: s.name, data: s.values, backgroundColor: s.values.map((_, i) => one || s.pointColors?.[i] || pal[i % pal.length]), borderColor: '#fff', borderWidth: ST.pieBorderPt * PX, pointStyle: 'rect' }];
    } else {
      const totals = pct ? labels.map((_, i) => series.reduce((a, s) => a + Math.abs(s.values[i] || 0), 0)) : null;
      // Excel column spacing: gap width and overlap are % of one bar's width
      const nBars = series.filter((s) => (s.type || t) === 'bar').length || 1;
      const sep = stacked ? 0 : Math.max(0, -ST.overlap) / 100, gap = (stacked ? ST.stackedGap : ST.gap) / 100;
      const unit = stacked ? 1 : nBars * (1 + sep);
      const barSpacing = { categoryPercentage: unit / (unit + gap), barPercentage: stacked ? 1 : 1 / (1 + sep) };
      datasets = series.map((s, i) => {
        const color = s.color || pal[i % pal.length];
        const st = s.type;
        const base = { label: s.name, backgroundColor: color, borderColor: color };
        let data = pct ? s.values.map((v, j) => (totals[j] ? (100 * (v || 0)) / totals[j] : 0)) : s.values;
        const r = (ST.scatterMarker * PX) / 2;
        if (t === 'scatter') return { ...base, type: 'scatter', data: s.values.map((y, j) => ({ x: s.cats ? s.cats[j] : j + 1, y })), pointRadius: r, pointHoverRadius: r + 2, showLine: false, pointStyle: 'circle' };
        // Excel's default line chart has no markers
        if (st === 'line') return { ...base, type: 'line', data, fill: false, tension: 0, pointRadius: 0, pointHoverRadius: 4, borderWidth: lineW, borderCapStyle: 'round', order: 0, pointStyle: 'line' };
        if (st === 'area') return { ...base, type: 'line', data, fill: stacked ? (i === 0 ? 'origin' : '-1') : 'origin', pointRadius: 0, borderWidth: 0, tension: 0, order: 1, pointStyle: 'rect' };
        if (st === 'radar') return { ...base, type: 'radar', data, backgroundColor: 'transparent', pointRadius: 0, borderWidth: lineW, pointStyle: 'line' };
        return { ...base, type: 'bar', data, borderWidth: 0, order: 2, pointStyle: 'rect', ...barSpacing };
      });
    }
    // horizontal clustered bars: Excel puts the first series at the bottom of each group; Chart.js
    // places a group's bars top-down in dataset order, so draw the datasets in reverse
    const flipped = horizontal && !stacked;
    if (flipped) datasets.reverse();
    const ctype = pieLike ? t : t === 'scatter' ? 'scatter' : t === 'radar' ? 'radar' : t === 'line' || t === 'area' ? 'line' : 'bar';
    const side = model.legend === 'r' || model.legend === 'l' || model.legend === 'tr';
    // Excel lists a side legend bottom-up when the series are drawn bottom-up (stacked, horizontal);
    // the datasets of flipped charts are already in reverse
    const legendReversed = (side && !pieLike && (stacked || horizontal)) !== flipped;
    const legendPos = { r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right' }[model.legend] || null;
    const family = `"${wb?.defaultFont?.name || 'Malgun Gothic'}", "Malgun Gothic", Calibri, sans-serif`;
    const text = { family, size: ST.textPt * PX };
    const opts = {
      responsive: false, animation: false, maintainAspectRatio: false,
      color: ST.text,
      plugins: {
        legend: {
          display: !!legendPos, position: legendPos || 'right',
          reverse: legendReversed,
          // symbols as in Excel: a line for lines, a square for bars, areas and pie slices
          labels: { usePointStyle: true, pointStyle: pieLike ? 'rect' : undefined, boxWidth: 8, boxHeight: 8, color: ST.text, font: text, sort: (a, b) => (a.datasetIndex - b.datasetIndex) || (a.index - b.index) },
        },
        title: { display: !!model.title, text: model.title || '', font: { family, size: ST.titlePt * PX, weight: 'normal' }, color: ST.text },
        tooltip: { callbacks: pct ? { label: (c) => `${c.dataset.label}: ${c.parsed.y?.toFixed?.(1) ?? c.parsed.x?.toFixed?.(1)}%` } : {} },
      },
      layout: { padding: 8 },
    };
    if (!pieLike && t !== 'radar') {
      const axisTitle = (s) => ({ display: !!s, text: s || '', color: ST.text, font: { family, size: 10 * PX } });
      // value axis: light gridlines, no axis line; category axis: light axis line, no gridlines
      const valAxis = { stacked, beginAtZero: true, grid: { color: ST.line, lineWidth: 0.75 * PX, drawTicks: false }, border: { display: false }, ticks: { color: ST.text, font: text, padding: 6 }, title: axisTitle(horizontal ? model.xTitle : model.yTitle) };
      const catAxis = { stacked, grid: { display: false, drawTicks: false }, border: { display: true, color: ST.line, width: 0.75 * PX }, ticks: { color: ST.text, font: text, padding: 6 }, title: axisTitle(horizontal ? model.yTitle : model.xTitle) };
      if (pct) valAxis.max = 100;
      if (t === 'scatter') { catAxis.type = 'linear'; catAxis.grid = { color: ST.line, lineWidth: 0.75 * PX, drawTicks: false }; }
      opts.scales = horizontal ? { x: valAxis, y: catAxis } : { x: catAxis, y: valAxis };
      if (horizontal) { opts.indexAxis = 'y'; catAxis.reverse = true; }
    }
    if (t === 'radar') opts.scales = { r: { grid: { color: ST.line }, angleLines: { color: ST.line }, pointLabels: { color: ST.text, font: text }, ticks: { color: ST.text, font: text, backdropColor: 'transparent' } } };
    return { type: ctype, data: { labels, datasets }, options: opts };
  }

  // ------------------------------------------------------------ DOM
  layout() {
    const doc = this.app.doc;
    if (!doc || typeof Chart === 'undefined') return;
    const sh = doc.wb.sheet;
    const m = this.grid.metrics();
    if (this.sheetShown !== sh) {
      for (const [, inst] of this.instances) { inst.chart.destroy(); inst.el.remove(); }
      this.instances.clear();
      this.sheetShown = sh;
      this.selected = null;
    }
    // clip charts below the headers
    const z = m.z;
    const sc = this.grid.scroller;
    this.layer.style.clipPath = `polygon(${sc.scrollLeft + m.HW * z}px ${sc.scrollTop + m.HH * z}px, 100% ${sc.scrollTop + m.HH * z}px, 100% 100%, ${sc.scrollLeft + m.HW * z}px 100%)`;
    const seen = new Set();
    for (const model of sh.charts) {
      if (!model.id) model.id = newId('ch');
      seen.add(model.id);
      let inst = this.instances.get(model.id);
      if (!inst) inst = this.create(model);
      const a = model.anchor;
      const x = (m.HW + m.g.cols.pos(a.c) + (a.dx || 0)) * z, y = (m.HH + m.g.rows.pos(a.r) + (a.dy || 0)) * z;
      Object.assign(inst.el.style, { left: x + 'px', top: y + 'px', width: a.w * z + 'px', height: a.h * z + 'px' });
      inst.model = model;
      const sig = JSON.stringify(model) + '|' + this.app.doc.wb.version + '|' + z;
      if (inst.sig !== sig) {
        inst.sig = sig;
        const cfg = this.config(model, sh);
        // logical size only: Chart.js multiplies by devicePixelRatio itself (setting both draws at 1/dpr scale)
        inst.canvas.style.width = a.w * z + 'px'; inst.canvas.style.height = a.h * z + 'px';
        if (!inst.chart) { inst.canvas.width = a.w * z; inst.canvas.height = a.h * z; }
        if (inst.chart && inst.type === cfg.type) {
          inst.chart.data = cfg.data; inst.chart.options = cfg.options; inst.chart.resize(a.w * z, a.h * z); inst.chart.update('none');
        } else {
          inst.chart?.destroy();
          inst.chart = new Chart(inst.canvas, { ...cfg, options: { ...cfg.options, devicePixelRatio: window.devicePixelRatio || 1 } });
          inst.type = cfg.type;
        }
      }
      inst.el.classList.toggle('selected', this.selected === model.id);
    }
    for (const [id, inst] of [...this.instances]) if (!seen.has(id)) { inst.chart?.destroy(); inst.el.remove(); this.instances.delete(id); }
  }

  create(model) {
    const el = document.createElement('div');
    el.className = 'chart-box';
    el.innerHTML = '<canvas></canvas><div class="ch-handle nw" data-h="nw"></div><div class="ch-handle ne" data-h="ne"></div><div class="ch-handle sw" data-h="sw"></div><div class="ch-handle se" data-h="se"></div>';
    this.layer.append(el);
    const inst = { el, canvas: el.querySelector('canvas'), chart: null, sig: null, model };
    this.instances.set(model.id, inst);
    el.addEventListener('pointerdown', (e) => this.onDown(e, inst));
    el.addEventListener('dblclick', () => this.app.chartDialog(inst.model));
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault(); this.select(inst.model.id);
      this.app.menu(e.clientX, e.clientY, [
        ['차트 편집…', () => this.app.chartDialog(inst.model)],
        ['차트 삭제', () => this.remove(inst.model.id)],
        ['맨 앞으로 가져오기', () => this.reorder(inst.model.id, 1)],
        ['맨 뒤로 보내기', () => this.reorder(inst.model.id, -1)],
      ]);
    });
    return inst;
  }

  select(id) { this.selected = id; this.app.editor?.active && this.app.editor.commit(); this.layout(); }
  deselect() { if (this.selected) { this.selected = null; this.layout(); } }

  onDown(e, inst) {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    this.select(inst.model.id);
    this.app.focusGrid();
    const handle = e.target.dataset.h;
    const sh = this.app.doc.wb.sheet;
    const z = this.grid.zoom;
    const g = geom(sh);
    const a0 = { ...inst.model.anchor };
    const x0 = g.cols.pos(a0.c) + (a0.dx || 0), y0 = g.rows.pos(a0.r) + (a0.dy || 0);
    const sx = e.clientX, sy = e.clientY;
    inst.el.setPointerCapture(e.pointerId);
    let next = null;
    const move = (ev) => {
      const dx = (ev.clientX - sx) / z, dy = (ev.clientY - sy) / z;
      let x = x0, y = y0, w = a0.w, h = a0.h;
      if (!handle) { x = Math.max(0, x0 + dx); y = Math.max(0, y0 + dy); }
      else {
        if (handle.includes('e')) w = Math.max(80, a0.w + dx);
        if (handle.includes('s')) h = Math.max(60, a0.h + dy);
        if (handle.includes('w')) { w = Math.max(80, a0.w - dx); x = x0 + a0.w - w; }
        if (handle.includes('n')) { h = Math.max(60, a0.h - dy); y = y0 + a0.h - h; }
      }
      const c = g.cols.at(x), r = g.rows.at(y);
      next = { c, r, dx: Math.round(x - g.cols.pos(c)), dy: Math.round(y - g.rows.pos(r)), w: Math.round(w), h: Math.round(h) };
      inst.model.anchor = next; // live (committed on release)
      this.layout();
    };
    const up = () => {
      inst.el.removeEventListener('pointermove', move); inst.el.removeEventListener('pointerup', up);
      if (!next) return;
      inst.model.anchor = a0;
      this.update(inst.model.id, (m) => { m.anchor = next; }, '차트 이동');
    };
    inst.el.addEventListener('pointermove', move);
    inst.el.addEventListener('pointerup', up);
  }

  update(id, fn, label = '차트') {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin(label); wb.touchMeta(sh);
    const m = sh.charts.find((c) => c.id === id);
    if (m) { fn(m); if (label !== '차트 이동') m.dirty = true; }
    wb.commit();
    this.layout();
  }

  remove(id) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin('차트 삭제'); wb.touchMeta(sh);
    sh.charts = sh.charts.filter((c) => c.id !== id);
    wb.commit();
    this.selected = null;
    this.layout();
  }

  reorder(id, dir) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin('차트 순서'); wb.touchMeta(sh);
    const i = sh.charts.findIndex((c) => c.id === id);
    const [c] = sh.charts.splice(i, 1);
    if (dir > 0) sh.charts.push(c); else sh.charts.unshift(c);
    wb.commit();
    for (const inst of this.instances.values()) inst.el.remove();
    for (const m of sh.charts) { const inst = this.instances.get(m.id); if (inst) this.layer.append(inst.el); }
  }

  /** Build a chart model from a data range (first row = series names, first col = categories when textual). */
  static fromRange(wb, sh, range, type = 'bar', byRows = null) {
    const val = (r, c) => wb.value(sh, r, c);
    const g = { ...range, r2: Math.min(range.r2, Math.max(sh.maxRow, range.r1)), c2: Math.min(range.c2, Math.max(sh.maxCol, range.c1)) };
    const rows = g.r2 - g.r1 + 1, cols = g.c2 - g.c1 + 1;
    let headerRow = false, headerCol = false;
    for (let c = g.c1; c <= g.c2; c++) if (typeof val(g.r1, c) === 'string') headerRow = true;
    for (let r = g.r1 + (headerRow ? 1 : 0); r <= g.r2; r++) if (typeof val(r, g.c1) === 'string' || (sh.get(r, g.c1)?.s?.nf && /[ymd]/.test(sh.get(r, g.c1).s.nf))) headerCol = true;
    if (cols === 1) headerCol = false;
    if (rows === 1) headerRow = false;
    if (headerRow && headerCol === false && typeof val(g.r1, g.c1) !== 'string' && cols > 1) { /* numeric corner */ }
    if (byRows === null) byRows = (cols - (headerCol ? 1 : 0)) > (rows - (headerRow ? 1 : 0)) && type !== 'scatter';
    const q = (r1, c1, r2, c2) => `${quoteSheetName(sh.name)}!${absRangeA1({ r1, c1, r2, c2 })}`;
    const dr1 = g.r1 + (headerRow ? 1 : 0), dc1 = g.c1 + (headerCol ? 1 : 0);
    const series = [];
    if (!byRows) {
      const cat = headerCol ? q(dr1, g.c1, g.r2, g.c1) : null;
      for (let c = dc1; c <= g.c2; c++) series.push({ name: null, nameRef: headerRow ? q(g.r1, c, g.r1, c) : null, cat, val: q(dr1, c, g.r2, c) });
    } else {
      const cat = headerRow ? q(g.r1, dc1, g.r1, g.c2) : null;
      for (let r = dr1; r <= g.r2; r++) series.push({ name: null, nameRef: headerCol ? q(r, g.c1, r, g.c1) : null, cat, val: q(r, dc1, r, g.c2) });
    }
    if (type === 'scatter' && series.length > 1 && !headerCol) {
      const x = series.shift();
      for (const s of series) s.cat = x.val;
    }
    const geomv = geom(sh);
    const anchorC = Math.min(MAX_COLS - 10, g.c2 + 2), anchorR = g.r1;
    return {
      id: newId('ch'), type, title: null, legend: type === 'pie' || type === 'doughnut' || series.length > 1 ? 'r' : null,
      barDir: 'col', grouping: type === 'line' || type === 'area' ? 'standard' : 'clustered',
      series, anchor: { c: anchorC, r: anchorR, dx: 0, dy: 0, w: 480, h: 288 }, dirty: true,
      _range: rangeA1(g),
    };
  }
}
