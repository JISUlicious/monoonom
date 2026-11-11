// Change history panel: every change to the workbook since it was opened (the server's history),
// newest first, with who made it (the agent, you in this window, another window), what it touched,
// and an Undo button per change, so the user can take back an agent's change without asking it.
// Undo uses the server's batch undo (the same as `xl.py undo`): it is recorded as a new change, shown
// in every window, and refused (then asked) when the same cells changed again later.
import { api } from './api.js';
import { t } from './i18n.js';
import * as D from './dialogs.js';

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

export class HistoryPanel {
  constructor(app) {
    this.app = app;
    this.filter = 'all'; // all | agent | user
    this.el = null;
    this.timer = 0;
    this.items = [];
  }

  get open() { return !!this.el && !this.el.hidden; }

  toggle(show = !this.open) {
    if (!this.el) this.build();
    this.el.hidden = !show;
    document.body.classList.toggle('history-open', show);
    if (show) { this.place(); this.refresh(); }
  }

  /** Beside the grid: from the top of the grid to the bottom of the sheet tabs. */
  place() {
    const g = this.app.gridHost.getBoundingClientRect(), s = document.getElementById('sheetbar').getBoundingClientRect();
    Object.assign(this.el.style, { top: g.top + 'px', height: Math.max(120, s.bottom - g.top) + 'px' });
  }

  build() {
    const el = this.el = document.createElement('aside');
    el.id = 'history-panel';
    el.className = 'history-panel';
    el.hidden = true;
    el.innerHTML = `<div class="hp-head"><b>${t('Change History')}</b>
      <select class="hp-filter" aria-label="${t('Show')}"><option value="all">${t('All changes')}</option><option value="agent">${t('Agent')}</option><option value="user">${t('People')}</option></select>
      <button class="hp-close" title="${t('Close')}">✕</button></div>
      <div class="hp-list" role="list"></div>
      <div class="hp-foot">${t('Since the workbook was opened. Undo takes back one change; it is recorded as a new change.')}</div>`;
    document.body.append(el);
    window.addEventListener('resize', () => { if (this.open) this.place(); });
    el.addEventListener('pointerdown', (e) => e.stopPropagation());
    el.querySelector('.hp-close').addEventListener('click', () => this.toggle(false));
    el.querySelector('.hp-filter').addEventListener('change', (e) => { this.filter = e.target.value; this.render(); });
    el.querySelector('.hp-list').addEventListener('click', (e) => {
      const u = e.target.closest('[data-undo]');
      if (u) return this.undo(u.dataset.undo);
      const g = e.target.closest('[data-goto]');
      if (g) this.app.goToRef(g.dataset.goto);
    });
  }

  /** Called on every change the window learns of (its own echoes too); refreshes while open. */
  onChange(doc) {
    if (!this.open || doc !== this.app.doc) return;
    clearTimeout(this.timer);
    this.timer = setTimeout(() => this.refresh(), 250);
  }

  /** Another workbook tab became active. */
  onDocSwitch() { if (this.open) this.refresh(); }

  async refresh() {
    const doc = this.app.doc;
    if (!doc) { this.items = []; return this.render(); }
    try {
      const res = await api.changes(doc.wb.docId);
      if (doc !== this.app.doc) return;
      this.items = res.changes || [];
    } catch (e) { this.items = []; this.error = e.message; }
    this.render();
  }

  who(c) {
    if (c.actor === 'agent') return ['agent', t('Agent')];
    if (c.actor === 'system') return ['system', t('Automatic')];
    return c.clientId === this.app.clientId ? ['me', t('You')] : ['other', t('Another window')];
  }

  render() {
    if (!this.el) return;
    const list = this.el.querySelector('.hp-list');
    const undone = new Set(this.items.map((c) => c.undoOf).filter(Boolean));
    const shown = this.items.filter((c) => this.filter === 'all' || (this.filter === 'agent' ? c.actor === 'agent' : c.actor === 'user')).reverse();
    if (!shown.length) { list.innerHTML = `<div class="hp-empty">${esc(this.error || t('No changes yet.'))}</div>`; this.error = null; return; }
    list.innerHTML = shown.map((c) => {
      const [k, name] = this.who(c);
      const where = (c.cells || []).map((r) => {
        const m = / \((\d+) cells?\)$/.exec(r), ref = m ? r.slice(0, m.index) : r;
        const n = m ? ` <small>${t('{n} cells', { n: m[1] })}</small>` : '';
        return `<a class="hp-ref" data-goto="${esc(ref)}" title="${t('Go to {ref}', { ref: esc(ref) })}">${esc(ref)}</a>${n}`;
      }).concat((c.meta || []).map((m) => `<span class="hp-meta">${esc(m)}</span>`)).join(' ');
      const isUndone = undone.has(c.batchId);
      const btn = c.undoOf || isUndone ? '' : `<button class="hp-undo" data-undo="${esc(c.batchId)}" title="${t('Undo this change')}">↶ ${t('Undo')}</button>`;
      return `<div class="hp-item${isUndone ? ' undone' : ''}" role="listitem" data-batch="${esc(c.batchId)}">
        <div class="hp-line"><span class="hp-who ${k}">${esc(name)}</span><span class="hp-time">${esc(c.time)}</span>${btn}</div>
        <div class="hp-label">${esc(c.label || t('Edit'))}${isUndone ? ` <em>(${t('undone')})</em>` : ''}</div>
        ${where ? `<div class="hp-where">${where}</div>` : ''}</div>`;
    }).join('');
  }

  async undo(batchId, force = false) {
    const doc = this.app.doc;
    if (this.app.editor?.active) this.app.editor.commit();
    try {
      await api.undo(doc.wb.docId, { batchId, actor: 'user', clientId: this.app.clientId, force });
      const c = this.items.find((x) => x.batchId === batchId);
      this.app.setStatus(t('Undone: {label}', { label: c?.label || '' }));
      this.refresh();
    } catch (e) {
      if (e.status === 409 && !force) {
        const ok = await D.confirmBox(t('{error}\nUndo it anyway? The later changes to those cells are replaced.', { error: e.message }), { yes: t('Undo'), no: t('Cancel') });
        if (ok === 'yes') return this.undo(batchId, true);
        return;
      }
      this.app.toast(e.message, 'error');
    }
  }
}
