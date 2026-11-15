/* wordPortFake.js — מסמך דמה בזיכרון עם אותו חוזה כמו wordPort (ראו core.js).
   משמש לבדיקות Node ולמצב ההדגמה בדפדפן.
   כל תיבה זוכרת מאיפה באה (origin): {pid, idx} מהמקור, {cid} מכתיבה של החלונית, או 'user'.
   כך הבדיקות יודעות אם החלונית כתבה במקום הנכון. מחיקות במעקב נשמרות כתיבות מחוקות (del),
   ו-getSimanParagraphs מחזיר רק את הטקסט הנקי, כמו getReviewedText("Current"). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.WordPortFake = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const AUTHOR = 'y.m. b';
  const words = (s) => String(s).split(/\s+/).filter(Boolean);

  class FakeDoc {
    constructor() { this.paras = []; this.tracking = true; this.revs = []; this.comments = []; this.nextKey = 1; this.violations = []; this.P = null; this.sel = null; this.cursorCbs = []; }

    static fromPackage(pkg) {
      const d = new FakeDoc();
      for (const p of pkg.paragraphs)
        d.paras.push({ key: 'k' + d.nextKey++, style: p.style, toks: p.orig_tokens.map((t, idx) => ({ t, del: false, origin: { pid: p.pid, idx } })) });
      return d;
    }
    setPackage(P) { this.P = P; }       // P = core.loadPackage(pkg), לבדיקת מקום הכתיבה

    // ───── עזרים
    _para(key) { const p = this.paras.find((x) => x.key === key); if (!p) throw new Error('אין פסקה ' + key); return p; }
    _clean(p) { return p.toks.filter((x) => !x.del); }
    _ix(p, i) {                                    // אינדקס נקי -> אינדקס פנימי
      let n = -1;
      for (let k = 0; k < p.toks.length; k++) if (!p.toks[k].del && ++n === i) return k;
      throw new Error('תיבה ' + i + ' לא קיימת בפסקה ' + p.key);
    }
    _rev(kind, text) { this.revs.push({ kind, text, author: AUTHOR, time: Date.now() }); }
    text(key) { return this._clean(this._para(key)).map((x) => x.t).join(' '); }
    dump() { return this.paras.map((p) => ({ style: p.style, text: this._clean(p).map((x) => x.t).join(' ') })); }

    // ───── בדיקת מקום הכתיבה (רק בבדיקות)
    _okTok(tok, tag, kind) {
      if (!this.P || !tag) return true;
      const o = tok.origin;
      if (o && o.cid === tag) return true;
      const c = this.P.clusters.get(tag);
      if (c) {
        if (o && o.pid === c.pid && o.idx >= c.from && o.idx <= c.to) return true;
        return false;
      }
      const it = this.P.items.get(tag);
      if (it) {
        if (o && o.pid === it.pid && o.idx === it.from) return true;
        if (o && o.cid) { const c2 = this.P.clusters.get(o.cid); if (c2 && c2.pid === it.pid && c2.from <= it.from && it.from <= c2.to) return true; }
        return false;
      }
      return true;
    }
    _check(edit, toks) {
      for (const t of toks) if (!this._okTok(t, edit.tag, edit.kind)) {
        this.violations.push({ edit, token: t.t, origin: t.origin });
        return false;
      }
      return true;
    }

    // ───── החוזה
    async getTracking() { return this.tracking; }
    async setTrackingOn() { if (!this.lockTracking) this.tracking = true; return this.tracking; }
    async getSimanParagraphs() { return this.paras.map((p) => ({ key: p.key, style: p.style, tokens: this._clean(p).map((x) => x.t) })); }
    async select(key, from, to) { this.sel = { key, from, to }; }
    async selectParagraph(key) { this.sel = { key }; }
    async insertComment(key, from, to, text) { this.comments.push({ key, from, to, text, author: AUTHOR }); }
    onCursorMoved(cb) { this.cursorCbs.push(cb); }
    async getCursorParagraph() { return this.sel ? this.sel.key : null; }

    async applyEdit(e) {
      if (!this.tracking) throw new Error('כתיבה בלי מעקב שינויים');
      const p = this._para(e.key);
      const mk = (t) => ({ t, del: false, origin: { cid: e.tag } });
      switch (e.kind) {
        case 'tokenEdit': {
          const k = this._ix(p, e.token), tok = p.toks[k];
          this._check(e, [tok]);
          tok.t = tok.t.slice(0, e.start) + e.text + tok.t.slice(e.start + e.len);
          if (!tok.t) tok.del = true;
          else if (/\s/.test(tok.t)) p.toks.splice(k, 1, ...words(tok.t).map((t) => ({ t, del: false, origin: tok.origin })));   // Word מחלק מחדש לתיבות
          this._rev('chars', e.text); return true;
        }
        case 'insertAfter': case 'insertBefore': {
          const k = this._ix(p, e.token);
          this._check(e, [p.toks[k]]);
          const nt = words(e.text).map(mk);
          p.toks.splice(e.kind === 'insertAfter' ? k + 1 : k, 0, ...nt);
          this._rev('ins', e.text); if (e.runs) this._rev('style', JSON.stringify(e.runs)); return true;
        }
        case 'replaceTokens': case 'deleteTokens': {
          const ks = []; for (let i = e.from; i <= e.to; i++) ks.push(this._ix(p, i));
          this._check(e, ks.map((k) => p.toks[k]));
          ks.forEach((k) => { p.toks[k].del = true; });
          if (e.kind === 'replaceTokens') p.toks.splice(ks[ks.length - 1] + 1, 0, ...words(e.text).map(mk));
          this._rev(e.kind === 'replaceTokens' ? 'replace' : 'del', e.text || ''); return true;
        }
        case 'split': {
          const k = this._ix(p, e.token);
          this._check(e, [p.toks[k]]);
          if (k === 0) throw new Error('פיצול בתחילת פסקה');
          const np = { key: 'k' + this.nextKey++, style: p.style, toks: p.toks.splice(k) };
          this.paras.splice(this.paras.indexOf(p) + 1, 0, np);
          this._rev('para', '¶'); return true;
        }
        case 'insertParagraphBefore': {
          const first = this._clean(p)[0];
          if (first) this._check(e, [first]);
          const np = { key: 'k' + this.nextKey++, style: e.style, toks: words(e.text).map(mk) };
          this.paras.splice(this.paras.indexOf(p), 0, np);
          this._rev('para', e.text); return true;
        }
      }
      throw new Error('פעולה לא מוכרת ' + e.kind);
    }

    // ───── "עריכות של יונתן" לבדיקות (במעקב, כמו ב-Word)
    userDeleteToken(key, i) { const p = this._para(key); const k = this._ix(p, i); p.toks[k].del = true; p.toks[k].touched = true; this._touchNeighbors(p, k); }
    userInsertToken(key, i, text) { const p = this._para(key); const k = i >= this._clean(p).length ? p.toks.length : this._ix(p, i); p.toks.splice(k, 0, { t: text, del: false, origin: 'user', touched: true }); this._touchNeighbors(p, k); }
    userEditToken(key, i, text) { const p = this._para(key); const k = this._ix(p, i); p.toks[k].t = text; p.toks[k].touched = true; }
    userSplit(key, i) { const p = this._para(key); const k = this._ix(p, i); const np = { key: 'k' + this.nextKey++, style: p.style, toks: p.toks.splice(k) }; this.paras.splice(this.paras.indexOf(p) + 1, 0, np); np.toks[0] && (np.toks[0].touched = true); p.toks.length && (p.toks[p.toks.length - 1].touched = true); }
    userAddParagraph(beforeKey, text, style) { const p = this._para(beforeKey); this.paras.splice(this.paras.indexOf(p), 0, { key: 'k' + this.nextKey++, style: style || p.style, toks: words(text).map((t) => ({ t, del: false, origin: 'user', touched: true })) }); }
    _touchNeighbors(p, k) { for (let d = -1; d <= 1; d++) { const x = p.toks[k + d]; if (x) x.touched = true; } }
  }

  return { FakeDoc, AUTHOR };
}));
