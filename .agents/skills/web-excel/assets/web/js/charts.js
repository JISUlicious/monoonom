// Chart layer: renders Excel chart models (same series refs as the xlsx) with Chart.js.
import { parseSheetRef, absRangeA1, rangeA1, MAX_ROWS, MAX_COLS } from './address.js';
import { quoteSheetName } from './formula.js';
import { formatValue } from './numfmt.js';
import { newId } from './model.js';
import { geom } from './geometry.js';

// Office theme accent colors (default Excel palette)
export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
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
      };
    });
    const labels = series.find((s) => s.cats)?.cats || (series[0] ? series[0].values.map((_, i) => String(i + 1)) : []);
    return { series, labels };
  }

  config(model, sheet) {
    const { series, labels } = this.dataOf(model, sheet);
    const t = model.type;
    const pieLike = t === 'pie' || t === 'doughnut';
    const horizontal = t === 'bar' && model.barDir === 'bar';
    const stacked = model.grouping === 'stacked' || model.grouping === 'percentStacked';
    const pct = model.grouping === 'percentStacked';
    let datasets;
    if (pieLike) {
      const s = series[0] || { values: [], name: '' };
      datasets = [{ label: s.name, data: s.values, backgroundColor: s.values.map((_, i) => PALETTE[i % PALETTE.length]), borderColor: '#fff', borderWidth: 1 }];
    } else {
      const totals = pct ? labels.map((_, i) => series.reduce((a, s) => a + Math.abs(s.values[i] || 0), 0)) : null;
      datasets = series.map((s, i) => {
        const color = s.color || PALETTE[i % PALETTE.length];
        const st = s.type;
        const base = { label: s.name, backgroundColor: color, borderColor: color };
        let data = pct ? s.values.map((v, j) => (totals[j] ? (100 * (v || 0)) / totals[j] : 0)) : s.values;
        if (t === 'scatter') return { ...base, type: 'scatter', data: s.values.map((y, j) => ({ x: s.cats ? s.cats[j] : j + 1, y })), pointRadius: 4, showLine: false };
        if (st === 'line') return { ...base, type: 'line', data, fill: false, tension: 0, pointRadius: 3, borderWidth: 2.25, order: 0 };
        if (st === 'area') return { ...base, type: 'line', data, fill: stacked ? (i === 0 ? 'origin' : '-1') : 'origin', backgroundColor: color + 'cc', pointRadius: 0, tension: 0, order: 1 };
        if (st === 'radar') return { ...base, type: 'radar', data, backgroundColor: color + '33', borderWidth: 2 };
        return { ...base, type: 'bar', data, order: 2 };
      });
    }
    const ctype = pieLike ? t : t === 'scatter' ? 'scatter' : t === 'radar' ? 'radar' : t === 'line' || t === 'area' ? 'line' : 'bar';
    const legendPos = { r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right' }[model.legend] || null;
    const font = { family: '"Malgun Gothic", Calibri, sans-serif' };
    const opts = {
      responsive: false, animation: false, maintainAspectRatio: false,
      plugins: {
        legend: { display: !!legendPos, position: legendPos || 'right', labels: { boxWidth: 10, font: { ...font, size: 11 } } },
        title: { display: !!model.title, text: model.title || '', font: { ...font, size: 14, weight: 'normal' }, color: '#595959' },
        tooltip: { callbacks: pct ? { label: (c) => `${c.dataset.label}: ${c.parsed.y?.toFixed?.(1) ?? c.parsed.x?.toFixed?.(1)}%` } : {} },
      },
      layout: { padding: 6 },
    };
    if (!pieLike && t !== 'radar') {
      const valAxis = { stacked, beginAtZero: true, grid: { color: '#d9d9d9' }, ticks: { font: { ...font, size: 10 } }, title: { display: !!(horizontal ? model.xTitle : model.yTitle), text: (horizontal ? model.xTitle : model.yTitle) || '' } };
      const catAxis = { stacked, grid: { display: false }, ticks: { font: { ...font, size: 10 } }, title: { display: !!(horizontal ? model.yTitle : model.xTitle), text: (horizontal ? model.yTitle : model.xTitle) || '' } };
      if (pct) valAxis.max = 100;
      if (t === 'scatter') { catAxis.type = 'linear'; catAxis.grid = { color: '#eee' }; }
      opts.scales = horizontal ? { x: valAxis, y: catAxis } : { x: catAxis, y: valAxis };
      if (horizontal) opts.indexAxis = 'y';
    }
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
        inst.canvas.width = a.w * z * (window.devicePixelRatio || 1); inst.canvas.height = a.h * z * (window.devicePixelRatio || 1);
        inst.canvas.style.width = a.w * z + 'px'; inst.canvas.style.height = a.h * z + 'px';
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
