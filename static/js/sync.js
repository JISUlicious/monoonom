// Multi-window support: windows of the same origin share open documents through BroadcastChannel.
// A workbook opened in several windows stays in sync (cell/meta/structure patches).
import { newId } from './model.js';

export class WindowSync {
  constructor(app) {
    this.app = app;
    this.winId = newId('w');
    this.peers = new Map(); // winId -> {title, docs:[{docId,name}], seen}
    this.applying = false;
    this.pendingFull = new Map();
    try { this.ch = new BroadcastChannel('web-excel'); } catch { this.ch = null; }
    if (!this.ch) return;
    this.ch.onmessage = (e) => this.onMessage(e.data);
    this.announce('hello');
    setInterval(() => { this.announce('here'); this.prune(); }, 4000);
    window.addEventListener('beforeunload', () => this.post({ type: 'bye' }));
  }

  post(msg) { if (this.ch) this.ch.postMessage({ ...msg, from: this.winId }); }

  announce(type = 'here') {
    this.post({ type, title: document.title, docs: this.app.docs.map((d) => ({ docId: d.wb.docId, name: d.wb.name })) });
  }

  prune() {
    const now = Date.now();
    let changed = false;
    for (const [id, p] of this.peers) if (now - p.seen > 10000) { this.peers.delete(id); changed = true; }
    if (changed) this.app.onWindowsChanged?.();
  }

  windows() {
    return [{ winId: this.winId, me: true, title: document.title, docs: this.app.docs.map((d) => ({ docId: d.wb.docId, name: d.wb.name })) },
      ...[...this.peers.entries()].map(([winId, p]) => ({ winId, ...p }))];
  }

  sharedWith(docId) { for (const p of this.peers.values()) if (p.docs.some((d) => d.docId === docId)) return true; return false; }

  /** Serialize a committed transaction (by current state) and broadcast. */
  publish(doc, entry) {
    if (!this.ch || this.applying || !this.sharedWith(doc.wb.docId)) return;
    const msg = { type: 'patch', docId: doc.wb.docId, cells: [], meta: [], structure: null };
    for (const ch of entry.cells) msg.cells.push([ch.sheet.id, ch.k, ch.sheet.cells.get(ch.k) || null]);
    for (const m of entry.meta) msg.meta.push([m.sheet.id, m.sheet.metaSnapshot()]);
    if (entry.structure) msg.structure = { order: doc.wb.sheets.map((s) => s.id) };
    this.post(msg);
  }

  publishSaved(doc) {
    this.post({ type: 'saved', docId: doc.wb.docId, path: doc.wb.path, name: doc.wb.name, format: doc.wb.format });
  }

  requestFull(docId) {
    return new Promise((resolve) => {
      this.pendingFull.set(docId, resolve);
      this.post({ type: 'need-full', docId });
      setTimeout(() => { if (this.pendingFull.get(docId) === resolve) { this.pendingFull.delete(docId); resolve(null); } }, 900);
    });
  }

  onMessage(m) {
    if (!m || m.from === this.winId) return;
    const app = this.app;
    switch (m.type) {
      case 'hello': case 'here': {
        const isNew = !this.peers.has(m.from);
        this.peers.set(m.from, { title: m.title, docs: m.docs || [], seen: Date.now() });
        if (m.type === 'hello') this.announce('here');
        if (isNew) app.onWindowsChanged?.();
        break;
      }
      case 'bye': this.peers.delete(m.from); app.onWindowsChanged?.(); break;
      case 'need-full': {
        const doc = app.docs.find((d) => d.wb.docId === m.docId);
        if (doc) this.post({ type: 'full', docId: m.docId, to: m.from, json: doc.wb.toJSON(), dirty: doc.wb.dirty, basePath: doc.wb.basePath });
        break;
      }
      case 'full': {
        if (m.to !== this.winId) break;
        const res = this.pendingFull.get(m.docId);
        if (res) { this.pendingFull.delete(m.docId); res(m); }
        else app.replaceFromFull?.(m);
        break;
      }
      case 'patch': {
        const doc = app.docs.find((d) => d.wb.docId === m.docId);
        if (!doc) break;
        const wb = doc.wb;
        const byId = new Map(wb.sheets.map((s) => [s.id, s]));
        const unknown = (m.structure?.order || []).some((id) => !byId.has(id)) || m.cells.some(([sid]) => !byId.has(sid)) || m.meta.some(([sid]) => !byId.has(sid));
        if (unknown) { this.requestFull(m.docId).then((full) => full && app.replaceFromFull(full)); break; }
        this.applying = true;
        try {
          if (m.structure) {
            const activeSheet = wb.sheet;
            wb.sheets = m.structure.order.map((id) => byId.get(id));
            wb.active = Math.max(0, wb.sheets.indexOf(activeSheet));
          }
          for (const [sid, meta] of m.meta) byId.get(sid).applyMeta(meta);
          const touched = new Set();
          for (const [sid, k, cell] of m.cells) {
            const sh = byId.get(sid);
            if (cell) sh.cells.set(k, cell); else sh.cells.delete(k);
            touched.add(sh);
          }
          for (const sh of touched) sh.updateExtent();
          wb.dirty = true;
          wb.version++;
          wb.recalc();
          app.onRemoteChange(doc);
        } finally { this.applying = false; }
        break;
      }
      case 'saved': {
        const doc = app.docs.find((d) => d.wb.docId === m.docId);
        if (doc) { Object.assign(doc.wb, { path: m.path, name: m.name, format: m.format, basePath: m.path, dirty: false }); app.renderTabs(); }
        break;
      }
      default: break;
    }
  }
}
