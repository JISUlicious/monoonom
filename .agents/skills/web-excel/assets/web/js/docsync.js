// Client side of the server-owned document (P1).
// Local edits are applied immediately (optimistic), sent as ops, and tracked as "pending" until their
// echo arrives on the SSE stream. Remote ops never overwrite cells that still have pending local
// writes: the server will apply ours after theirs, so ours wins (last writer wins, server order).
import { api } from './api.js';
import { keyRow, keyCol, rangeA1, a1 } from './address.js';
import { sheetMetaJSON, sheetJSON, cellJSON, newId } from './model.js';

export class DocSync {
  constructor(app, doc, version) {
    this.app = app;
    this.doc = doc;
    this.version = version;
    this.pending = new Map();      // batchId -> {cells:[key], meta:[sid], structure:bool}
    this.pendCells = new Map();    // `${sid}:${k}` -> count
    this.pendMeta = new Map();     // sid -> count
    this.pendStruct = 0;
    this.known = new Set(doc.wb.sheets.map((s) => s.id));
    this.es = null;
    this.closed = false;
    this.lastCalc = new Map();
    this.calcTimer = 0;
    this.calcFull = true;
  }

  get id() { return this.doc.wb.docId; }

  /** subscribe this document on the window's shared event stream (replays ops after `version`) */
  connect() {
    if (this.closed) return;
    this.app.stream.subscribe(this);
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    clearTimeout(this.calcTimer);
    this.app.stream.unsubscribe(this);
  }

  // ------------------------------------------------------------ outgoing
  async send(entry) {
    const wb = this.doc.wb;
    const ops = [];
    const batch = { cells: [], meta: [], structure: false };
    const bySheet = new Map();
    for (const ch of entry.cells) {
      if (!bySheet.has(ch.sheet)) bySheet.set(ch.sheet, []);
      bySheet.get(ch.sheet).push(ch.k);
    }
    if (entry.structure) {
      const add = {};
      for (const s of wb.sheets) if (!this.known.has(s.id)) { add[s.id] = sheetJSON(s); this.known.add(s.id); }
      ops.push({ op: 'sheets', order: wb.sheets.map((s) => s.id), add });
      batch.structure = true;
    }
    for (const [sheet, keys] of bySheet) {
      if (!wb.sheets.includes(sheet)) continue; // sheet deleted in the same batch
      const uniq = [...new Set(keys)];
      ops.push({ op: 'setCells', sheet: sheet.id, cells: uniq.map((k) => [keyRow(k), keyCol(k), cellJSON(sheet.cells.get(k))]) });
      for (const k of uniq) batch.cells.push(`${sheet.id}:${k}`);
    }
    for (const m of entry.meta) {
      if (!wb.sheets.includes(m.sheet)) continue;
      ops.push({ op: 'setSheetMeta', sheet: m.sheet.id, meta: sheetMetaJSON(m.sheet) });
      batch.meta.push(m.sheet.id);
    }
    if (entry.meta.length || entry.structure) ops.push({ op: 'setWorkbook', props: { definedNames: wb.definedNames } });
    if (!ops.length) return;
    const batchId = newId('b');
    this.track(batchId, batch, +1);
    try {
      await api.ops(this.id, {
        ops, batchId, baseVersion: this.version, clientId: this.app.clientId, actor: 'user',
        label: entry.label, structural: !!entry.structural,
        transformable: entry.structural ? !!entry.transformable : undefined, shifts: entry.shifts?.length ? entry.shifts : undefined,
      });
    } catch (e) {
      this.track(batchId, batch, -1);
      if (e.status === 409) await this.resync(`${e.message} 마지막 작업("${entry.label}")은 반영되지 않았습니다.`);
      else this.app.toast(`서버에 변경을 보내지 못했습니다: ${e.message}`, 'error', 8000);
    }
  }

  track(batchId, batch, d) {
    const bump = (map, k) => { const n = (map.get(k) || 0) + d; if (n > 0) map.set(k, n); else map.delete(k); };
    for (const k of batch.cells) bump(this.pendCells, k);
    for (const s of batch.meta) bump(this.pendMeta, s);
    if (batch.structure) this.pendStruct = Math.max(0, this.pendStruct + d);
    if (d > 0) this.pending.set(batchId, batch); else this.pending.delete(batchId);
  }

  // ------------------------------------------------------------ incoming
  onOps(ev) {
    if (ev.v <= this.version) return; // replayed after reconnect
    this.version = ev.v;
    if (ev.clientId && ev.clientId === this.app.clientId) {
      const batch = this.pending.get(ev.batchId);
      if (batch) this.track(ev.batchId, batch, -1);
      // the server moved our edit through someone else's row/column insert/delete
      if (ev.transformed) this.needResync = true;
      this.maybeResync();
      return;
    }
    // a remote structural change while our edits are in flight: our local cells are at old
    // coordinates; once our echoes are back, reload the (correct) server state quietly
    if (ev.structural && this.pending.size) this.needResync = true;
    const wb = this.doc.wb;
    for (const op of ev.ops) if (op.op === 'sheets') for (const id of op.order) this.known.add(id);
    const skip = (kind, sid, k) => (kind === 'cell' ? this.pendCells.has(`${sid}:${k}`) : kind === 'meta' ? this.pendMeta.has(sid) : this.pendStruct > 0);
    const boxes = wb.applyRemote(ev.ops, skip);
    if (ev.actor !== 'system') wb.dirty = true;
    this.app.onRemoteChange(this.doc, ev, boxes);
    this.scheduleCalc();
  }

  maybeResync() {
    if (!this.needResync || this.pending.size) return;
    if (this.app.editor?.active && this.app.doc === this.doc) { setTimeout(() => this.maybeResync(), 500); return; }
    this.needResync = false;
    this.resync(null);
  }

  onSaved(info) {
    Object.assign(this.doc.wb, { path: info.path, name: info.name, format: info.format, basePath: info.path, dirty: false });
    for (const s of this.doc.wb.sheets) s.origName = s.name;
    this.app.renderTabs();
  }

  async resync(msg) {
    const snap = await api.snapshot(this.id);
    this.app.replaceDoc(this.doc, snap);
    this.pending.clear(); this.pendCells.clear(); this.pendMeta.clear(); this.pendStruct = 0;
    this.version = snap.version;
    this.known = new Set(this.doc.wb.sheets.map((s) => s.id));
    this.calcFull = true; this.lastCalc.clear();
    this.connect();
    this.scheduleCalc();
    if (msg) this.app.toast(msg, 'warn', 7000);
  }

  // ------------------------------------------------------------ presence & calc reports
  presence(sheet, view) {
    clearTimeout(this.presTimer);
    this.presTimer = setTimeout(() => {
      api.presence(this.id, { clientId: this.app.clientId, actor: 'user', sheet: sheet.id, sheetName: sheet.name,
        sel: rangeA1(view.sel), active: a1(view.active.r, view.active.c) }).catch(() => {});
    }, 250);
  }

  /** Report formula results so the agent can read computed values (the browser is a calc source). */
  scheduleCalc() {
    clearTimeout(this.calcTimer);
    this.calcTimer = setTimeout(() => this.reportCalc(), 400);
  }

  async reportCalc() {
    const wb = this.doc.wb;
    const values = [];
    const seen = new Set();
    for (const sh of wb.sheets) {
      for (const [k, cell] of sh.cells) {
        if (!cell.f) continue;
        const id = `${sh.id}:${k}`;
        seen.add(id);
        let v = wb.value(sh, keyRow(k), keyCol(k));
        if (v && typeof v === 'object') v = { err: v.err };
        const j = JSON.stringify(v ?? null);
        if (!this.calcFull && this.lastCalc.get(id) === j) continue;
        this.lastCalc.set(id, j);
        values.push([sh.id, keyRow(k), keyCol(k), v ?? null]);
      }
    }
    for (const id of [...this.lastCalc.keys()]) if (!seen.has(id)) this.lastCalc.delete(id);
    if (!values.length && !this.calcFull) {
      return api.calc(this.id, { version: this.version, values: [], source: 'browser' }).catch(() => {});
    }
    const full = this.calcFull;
    this.calcFull = false;
    try { await api.calc(this.id, { version: this.version, values, source: 'browser', full }); } catch { this.calcFull = true; }
  }
}


/**
 * One SSE connection per window (browsers cap HTTP/1.1 connections per host at ~6 across tabs).
 * Documents are subscribed over POST; events carry `docId` and are dispatched to their DocSync.
 */
export class EventStream {
  constructor(app) {
    this.app = app;
    this.syncs = new Map(); // docId -> DocSync
    this.ready = false;
    this.open();
  }

  /** EventSource cannot send headers: get a single-use ticket first (the server also accepts the
   * page's cookie). */
  async open() {
    this.ready = false;
    let ticket = '';
    try { ticket = (await api.streamTicket()).ticket; } catch (e) {
      this.app.onAuthError?.(e);
      setTimeout(() => this.open(), e.status === 401 ? 5000 : 1200);
      return;
    }
    const es = new EventSource(`/api/stream?client=${encodeURIComponent(this.app.clientId)}&ticket=${encodeURIComponent(ticket)}`);
    this.es = es;
    const on = (name, fn) => es.addEventListener(name, (e) => fn(JSON.parse(e.data)));
    on('ready', (d) => {
      this.ready = true;
      for (const sync of this.syncs.values()) this.post(sync);
      this.app.onDocsList?.(d.docs);
    });
    on('ops', (d) => this.syncs.get(d.docId)?.onOps(d));
    on('saved', (d) => this.syncs.get(d.docId)?.onSaved(d));
    on('resync', (d) => this.syncs.get(d.docId)?.resync('서버 기록이 앞서 있어 다시 불러왔습니다.'));
    on('presence', (d) => this.app.onPresence?.(d));
    on('goto', (d) => this.app.onGoto?.(d));
    on('view', (d) => this.app.onView?.(d));
    on('opened', (d) => this.app.onDocOpened?.(d));
    on('focus', (d) => this.app.onDocFocus?.(d));
    on('closed', (d) => {
      const sync = this.syncs.get(d.docId);
      if (sync) { sync.close(); this.app.onDocClosedRemotely(sync.doc); }
    });
    es.onerror = () => { es.close(); this.ready = false; setTimeout(() => this.open(), 1200); };
  }

  post(sync) {
    api.subscribe({ client: this.app.clientId, docId: sync.id, since: sync.version }).catch((e) => {
      if (e.status === 404 && /스트림/.test(e.message)) setTimeout(() => this.ready && this.post(sync), 500);
    });
  }

  subscribe(sync) {
    this.syncs.set(sync.id, sync);
    if (this.ready) this.post(sync);
  }

  unsubscribe(sync) {
    if (this.syncs.get(sync.id) !== sync) return;
    this.syncs.delete(sync.id);
    api.unsubscribe({ client: this.app.clientId, docId: sync.id }).catch(() => {});
  }
}
