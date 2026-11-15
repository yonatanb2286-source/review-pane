/* wordMock.js — דמה מינימלי של Word.js (Office) לבדיקת wordPortWord ב-Node.
   כל פסקה היא רשימת תווים {c, del, add, cstyle, mine0}. טקסט raw כולל תווים מחוקים במעקב (כמו paragraph.text ב-Word).
   הכתיבות רשומות במעקב בשם ME. מחיקה של הוספה של ME עצמו מעלימה אותה (כמו ב-Word). */
'use strict';
const ME = 'y.m. b';
let uid = 1;
class Doc {
  constructor() { this.paras = []; this.mode = 'TrackAll'; this.sel = null; this.comments = []; this.ops = 0; }
  addPara(text, style, extra) { const p = { id: 'P' + uid++, style, chars: [...text].map((c) => ({ c, del: false, add: null, cstyle: null })) }; Object.assign(p, extra || {}); this.paras.push(p); return p; }
  raw(p) { return p.chars.map((x) => x.c).join(''); }
  clean(p) { return p.chars.filter((x) => !x.del).map((x) => x.c).join(''); }
  pix(p) { return this.paras.indexOf(p); }
}
class Rng {
  constructor(doc, p0, a, p1, b) { this.doc = doc; this.p0 = p0; this.a = a; this.p1 = p1 === undefined ? p0 : p1; this.b = b === undefined ? a : b; }
  _one() { if (this.p0 !== this.p1) throw new Error('mock: טווח רב-פסקאות'); return this.p0; }
  load() { if (this.p0 === this.p1) { this.text = this.doc.raw(this.p0).slice(this.a, this.b); this._st = this._style(); } return this; }
  _style() { const ch = this.p0.chars.slice(this.a, this.b); const s = new Set(ch.map((x) => x.cstyle)); if (s.size === 1) { const v = [...s][0]; return v || this.p0.style; } return ''; }
  set style(v) { if (this.__init) { const ch = this.p0.chars.slice(this.a, this.b); ch.forEach((x) => { x.cstyle = v; }); this.doc.ops++; } else this._st = v; }
  get style() { return this._st; }
  getRange(loc) {
    if (loc === 'Start') return mk(this.doc, this.p0, this.a);
    if (loc === 'End') return mk(this.doc, this.p1, this.b);
    return mk(this.doc, this.p0, this.a, this.p1, this.b);
  }
  expandTo(o) {
    const d = this.doc, ia = d.pix(this.p0), ib = d.pix(o.p0);
    const [s0, sa] = (ia < ib || (ia === ib && this.a <= o.a)) ? [this.p0, this.a] : [o.p0, o.a];
    const ja = d.pix(this.p1), jb = d.pix(o.p1);
    const [e1, eb] = (ja > jb || (ja === jb && this.b >= o.b)) ? [this.p1, this.b] : [o.p1, o.b];
    return mk(d, s0, sa, e1, eb);
  }
  get paragraphs() { const d = this.doc; const L = d.paras.slice(d.pix(this.p0), d.pix(this.p1) + 1); return coll(L.map((p) => new PP(d, p))); }
  getReviewedText() { const p = this._one(); return { value: p.chars.slice(this.a, this.b).filter((x) => !x.del).map((x) => x.c).join('') + (this.b >= p.chars.length ? '\r' : '') }; }
  getTrackedChanges() {
    const p = this._one(), out = [];
    let cur = null;
    for (let k = this.a; k < this.b; k++) {
      const x = p.chars[k], type = x.del ? 'Deleted' : (x.add ? 'Added' : null), au = x.del ? x.delBy : x.add;
      if (cur && cur.type === type && cur.author === au && type) { cur.text += x.c; continue; }
      cur = type ? { type, author: au, text: x.c, start: k } : null;
      if (cur) out.push(cur);
    }
    const d = this.doc;
    return coll(out.map((t) => ({ type: t.type, text: t.text, author: t.author, getRange: () => mk(d, p, t.start) })));
  }
  search(s) {
    const p = this._one(), raw = this.doc.raw(p).slice(this.a, this.b), q = s.replace(/\^\^/g, '^'), res = [];
    let i = 0; while ((i = raw.indexOf(q, i)) >= 0) { res.push(mk(this.doc, p, this.a + i, p, this.a + i + q.length)); i += q.length; }
    return coll(res);
  }
  insertText(t, loc) {
    const d = this.doc, p = this._one();
    if (d.mode === 'Off') throw new Error('mock: כתיבה בלי מעקב');
    if (loc === 'Replace') { const r = mk(d, p, this.a, p, this.b); const removed = r.delete(); return mk(d, p, this.a).insertText(t, 'Start', removed); }
    const pos = loc === 'End' || loc === 'After' ? this.b : this.a;
    if (t === '\n') {
      if (pos === 0) throw new Error('mock: פיצול בתחילת פסקה');
      const np = { id: 'P' + uid++, style: p.style, chars: p.chars.splice(pos) };
      d.paras.splice(d.pix(p) + 1, 0, np); d.ops++;
      return mk(d, np, 0);
    }
    const inh = pos > 0 ? p.chars[pos - 1].cstyle : null;
    p.chars.splice(pos, 0, ...[...t].map((c) => ({ c, del: false, add: ME, cstyle: inh })));
    d.ops++;
    return mk(d, p, pos, p, pos + t.length);
  }
  delete() {
    const p = this._one(); let removed = 0;
    for (let k = this.b - 1; k >= this.a; k--) {
      const x = p.chars[k];
      if (x.add === ME && !x.del) { if (x.mine0) this.doc.lostMine = true; p.chars.splice(k, 1); removed++; } else if (!x.del) { x.del = true; x.delBy = ME; }
    }
    this.doc.ops++; return removed;
  }
  select() { this.doc.sel = this; }
  insertComment(t) { this.doc.comments.push(t); }
}
function mk(doc, p0, a, p1, b) { const r = new Rng(doc, p0, a, p1, b); r.__init = true; return r; }
function coll(items) { return { items, load() { items.forEach((x) => x.load && x.load()); return this; }, getFirst() { return items[0]; } }; }
class PP {
  constructor(doc, p) { this.doc = doc; this.p = p; }
  load() { this.text = this.doc.raw(this.p); this._style = this.p.style; this.uniqueLocalId = this.p.id; this.isNullObject = false; return this; }
  get style() { return this._style; }
  set style(v) { this.p.style = v; this._style = v; this.doc.ops++; }
  _r() { return mk(this.doc, this.p, 0, this.p, this.p.chars.length); }
  getRange(loc) { const r = this._r(); if (loc === 'Start') return mk(this.doc, this.p, 0); if (loc === 'End') return mk(this.doc, this.p, this.p.chars.length); return r; }
  search(s, o) { return this._r().search(s, o); }
  insertParagraph(t, loc) {
    const d = this.doc, i = d.pix(this.p);
    const np = { id: 'P' + uid++, style: this.p.style, chars: [...t].map((c) => ({ c, del: false, add: ME, cstyle: null })) };
    d.paras.splice(loc === 'Before' ? i : i + 1, 0, np); d.ops++;
    return new PP(d, np).load();
  }
  select() { this.doc.sel = this.getRange('Start'); }
}
function install(doc) {
  const ctx = {
    document: {
      load() {}, get changeTrackingMode() { return doc.mode; }, set changeTrackingMode(v) { doc.mode = v; },
      body: { get paragraphs() { return coll(doc.paras.map((p) => new PP(doc, p))); }, getRange(loc) { const L = doc.paras[doc.paras.length - 1]; return mk(doc, L, L.chars.length); } },
      getParagraphByUniqueLocalId(id) { const p = doc.paras.find((x) => x.id === id); if (!p) throw new Error('ItemNotFound'); return new PP(doc, p).load(); },
      getSelection() { const s = doc.sel; return { paragraphs: coll([new PP(doc, s ? s.p0 : doc.paras[0])]) }; },
    },
    sync: async () => {},
  };
  global.Word = { run: async (fn) => fn(ctx), ChangeTrackingMode: { trackAll: 'TrackAll' } };
  global.Office = { context: { document: { addHandlerAsync() {} } }, EventType: { DocumentSelectionChanged: 'sel' } };
  return ctx;
}
module.exports = { Doc, install, ME };
