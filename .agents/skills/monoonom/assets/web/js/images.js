// Pictures on a sheet: drawn over the grid like charts (same layer), selected, moved and resized with
// the mouse, deleted with Delete; pasted from the clipboard. The bytes live on the server (the
// document's media); a sheet's `images` hold {id, media, format, anchor {c, r, dx, dy, w, h}}.
import { api } from './api.js';
import { geom } from './geometry.js';
import { newId } from './model.js';
import { t } from './i18n.js';

export class ImageLayer {
  constructor(app, grid) {
    this.app = app;
    this.grid = grid;
    this.layer = grid.chartLayer;
    this.instances = new Map(); // image id -> {el, img, media}
    this.urls = new Map();      // media key -> object URL (fetched with the token header)
    this.selected = null;
    this.sheetShown = null;
  }

  /** Object URL of a picture's bytes (the page's token goes in a header, which an <img src> cannot send). */
  url(media) {
    if (!this.urls.has(media)) {
      this.urls.set(media, api.media(this.app.doc.wb.docId, media).then((b) => URL.createObjectURL(b)).catch(() => ''));
    }
    return this.urls.get(media);
  }

  layout() {
    const doc = this.app.doc;
    if (!doc) return;
    const sh = doc.wb.sheet;
    if (this.sheetShown !== sh) {
      for (const [, inst] of this.instances) inst.el.remove();
      this.instances.clear();
      this.sheetShown = sh;
      this.selected = null;
    }
    const m = this.grid.metrics(), z = m.z;
    const seen = new Set();
    for (const model of sh.images || []) {
      if (!model.id) model.id = newId('im');
      seen.add(model.id);
      let inst = this.instances.get(model.id);
      if (!inst) inst = this.create(model);
      if (inst.media !== model.media) {
        inst.media = model.media;
        this.url(model.media).then((u) => { if (inst.media === model.media) inst.img.src = u; });
      }
      const a = model.anchor;
      const x = (m.HW + m.g.cols.pos(a.c) + (a.dx || 0)) * z, y = (m.HH + m.g.rows.pos(a.r) + (a.dy || 0)) * z;
      Object.assign(inst.el.style, { left: x + 'px', top: y + 'px', width: a.w * z + 'px', height: a.h * z + 'px' });
      inst.model = model;
      inst.el.classList.toggle('selected', this.selected === model.id);
    }
    for (const [id, inst] of [...this.instances]) if (!seen.has(id)) { inst.el.remove(); this.instances.delete(id); }
  }

  create(model) {
    const el = document.createElement('div');
    el.className = 'chart-box image-box';
    el.innerHTML = '<img alt="" draggable="false"><div class="ch-handle nw" data-h="nw"></div><div class="ch-handle ne" data-h="ne"></div><div class="ch-handle sw" data-h="sw"></div><div class="ch-handle se" data-h="se"></div>';
    this.layer.append(el);
    const inst = { el, img: el.querySelector('img'), media: null, model };
    this.instances.set(model.id, inst);
    el.addEventListener('pointerdown', (e) => this.onDown(e, inst));
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault(); this.select(inst.model.id);
      this.app.menu(e.clientX, e.clientY, [
        [t('Delete Picture'), () => this.remove(inst.model.id)],
        [t('Reset Picture Size'), () => this.resetSize(inst.model.id)],
        [t('Bring to Front'), () => this.reorder(inst.model.id, 1)],
        [t('Send to Back'), () => this.reorder(inst.model.id, -1)],
      ]);
    });
    return inst;
  }

  select(id) {
    this.app.charts?.deselect();
    this.selected = id; this.app.editor?.active && this.app.editor.commit(); this.layout();
  }
  deselect() { if (this.selected) { this.selected = null; this.layout(); } }

  /** Drag to move; corner handles resize (Shift keeps the aspect ratio, as in Excel). */
  onDown(e, inst) {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    this.select(inst.model.id);
    this.app.focusGrid();
    const handle = e.target.dataset.h;
    const sh = this.app.doc.wb.sheet, z = this.grid.zoom, g = geom(sh);
    const a0 = { ...inst.model.anchor };
    const x0 = g.cols.pos(a0.c) + (a0.dx || 0), y0 = g.rows.pos(a0.r) + (a0.dy || 0);
    const sx = e.clientX, sy = e.clientY, ratio = a0.w / Math.max(1, a0.h);
    inst.el.setPointerCapture(e.pointerId);
    let next = null;
    const move = (ev) => {
      const dx = (ev.clientX - sx) / z, dy = (ev.clientY - sy) / z;
      let x = x0, y = y0, w = a0.w, h = a0.h;
      if (!handle) { x = Math.max(0, x0 + dx); y = Math.max(0, y0 + dy); }
      else {
        if (handle.includes('e')) w = Math.max(10, a0.w + dx);
        if (handle.includes('s')) h = Math.max(10, a0.h + dy);
        if (handle.includes('w')) w = Math.max(10, a0.w - dx);
        if (handle.includes('n')) h = Math.max(10, a0.h - dy);
        if (ev.shiftKey) { if (Math.abs(dx) >= Math.abs(dy)) h = w / ratio; else w = h * ratio; }
        if (handle.includes('w')) x = x0 + a0.w - w;
        if (handle.includes('n')) y = y0 + a0.h - h;
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
      this.update(inst.model.id, (m) => { m.anchor = next; }, handle ? t('Resize Picture') : t('Move Picture'));
    };
    inst.el.addEventListener('pointermove', move);
    inst.el.addEventListener('pointerup', up);
  }

  update(id, fn, label) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin(label); wb.touchMeta(sh);
    const m = sh.images.find((x) => x.id === id);
    if (m) fn(m);
    wb.commit();
    this.layout();
  }

  remove(id) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin(t('Delete Picture')); wb.touchMeta(sh);
    sh.images = sh.images.filter((x) => x.id !== id);
    wb.commit();
    this.selected = null;
    this.layout();
  }

  reorder(id, dir) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    wb.begin(t('Picture Order')); wb.touchMeta(sh);
    const i = sh.images.findIndex((x) => x.id === id);
    const [x] = sh.images.splice(i, 1);
    if (dir > 0) sh.images.push(x); else sh.images.unshift(x);
    wb.commit();
    for (const inst of this.instances.values()) inst.el.remove();
    for (const m of sh.images) { const inst = this.instances.get(m.id); if (inst) this.layer.append(inst.el); }
  }

  resetSize(id) {
    const inst = this.instances.get(id);
    const w = inst?.img.naturalWidth, h = inst?.img.naturalHeight;
    if (w && h) this.update(id, (m) => { m.anchor = { ...m.anchor, w, h }; }, t('Reset Picture Size'));
  }

  /** Paste a picture (a Blob from the clipboard) at the active cell: uploaded to the server, then added
   * to the sheet as one undoable change. Large pictures are shown at most 640 px wide (Excel shows
   * them at full size; the file keeps the full picture either way). */
  async paste(blob) {
    const doc = this.app.doc, wb = doc.wb;
    if (blob.size > 10 * 1024 * 1024) return this.app.toast(t('Pictures up to 10 MB can be pasted.'), 'warn');
    const data = await new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(String(fr.result).split(',')[1] || '');
      fr.onerror = () => reject(fr.error);
      fr.readAsDataURL(blob);
    });
    let info;
    try { info = await api.mediaAdd(wb.docId, { data }); } catch (e) { return this.app.toast(e.message, 'error'); }
    const k = Math.min(1, 640 / Math.max(1, info.w));
    const a = this.app.view().active, sh = wb.sheet;
    const im = { id: newId('im'), media: info.media, format: info.format, dirty: true,
      anchor: { c: a.c, r: a.r, dx: 0, dy: 0, w: Math.round(info.w * k), h: Math.round(info.h * k) } };
    wb.begin(t('Paste Picture')); wb.touchMeta(sh);
    sh.images = [...(sh.images || []), im];
    wb.commit();
    this.select(im.id);
    this.app.setStatus(t('Picture pasted ({w} × {h} px). Drag to move; drag a corner to resize (Shift keeps the shape).', { w: im.anchor.w, h: im.anchor.h }));
  }
}
