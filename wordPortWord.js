/* wordPortWord.js — שלב 4: ה-port האמיתי מול Word (Office.js), לפי החוזה שבראש core.js.
   כללי ברזל (תוכנית בנייה 5.1):
   - כותבים רק בתוך הסימן שנטען, ורק כשמעקב השינויים דולק.
   - לעולם לא מקבלים/דוחים שינויים במעקב, לא מוחקים הערות, לא משנים סגנון של פסקה קיימת.
   - לא מוחקים ולא מחליפים טקסט שהוא הוספה במעקב של יונתן עצמו (Word היה מעלים אותה בלי עקבות):
     במקרה כזה לא נכתב דבר, והחלונית מבקשת לעשות זאת ידנית.
   - כל כתיבה מאתרת את המקום מחדש (טקסט הפסקה + מחיקות במעקב), ובודקת שהטקסט במקום הוא בדיוק המצופה.
     אם משהו לא מתאים: לא נכתב דבר (מחזיר false, והסיבה ב-lastError).
   קואורדינטות: raw = paragraph.text (כולל טקסט מחוק במעקב). clean = getReviewedText("Current"). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.WordPortWord = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const SIMAN_STYLE = '1_סימן';
  const ME = 'y.m. b';
  const CHAR_MAP = { '8_מ"מ תו': '8_מ"מ' };            // ממצא שלב 0: השם בפועל הוא 8_מ"מ (סגנון מקושר)
  const ANCHOR = 24;
  const stripEnd = (s) => String(s == null ? '' : s).replace(/[\r\n\u0007]+$/, '');
  const normHead = (s) => String(s).replace(/[^\p{L}\p{N}\s]/gu, '').replace(/\s+/g, ' ').trim();

  // ───────── חלק טהור (נבדק ב-Node): מיפוי בין raw ל-clean
  /* dels: [{start, len}] בקואורדינטות raw. מחזיר {live:[bool], rawIdx:[clean->raw], toks:[{cs,ce}]} או זורק אם לא מתיישב */
  function buildMap(raw, clean, dels) {
    const live = new Array(raw.length).fill(true);
    for (const d of dels) for (let k = d.start; k < d.start + d.len && k < raw.length; k++) live[k] = false;
    const rawIdx = [];
    for (let k = 0; k < raw.length; k++) if (live[k]) rawIdx.push(k);
    const rebuilt = rawIdx.map((k) => raw[k]).join('');
    if (rebuilt !== clean) throw new Error('טקסט הפסקה לא מתיישב עם המחיקות במעקב');
    const toks = []; const re = /\S+/g; let m;
    while ((m = re.exec(clean))) toks.push({ cs: m.index, ce: m.index + m[0].length });
    return { raw, clean, live, rawIdx, toks };
  }
  /* נקודה ב-raw לפני תו נקי c ('before') או אחרי תו נקי c-1 ('after') */
  function rawPoint(M, c, side) {
    if (side === 'before') return c < M.rawIdx.length ? M.rawIdx[c] : M.raw.length;
    return c > 0 ? M.rawIdx[c - 1] + 1 : 0;
  }
  /* עוגן לחיפוש: רצף תווים חיים ב-raw שנגמר בנקודה P ('end') או מתחיל בה ('start'). בלי רווחים בקצה החיצוני. */
  function anchorAt(M, P, at) {
    let a = P, b = P;
    if (at === 'end') { while (a > 0 && M.live[a - 1] && P - a < ANCHOR) a--; while (a < P && /\s/.test(M.raw[a])) a++; }
    else { while (b < M.raw.length && M.live[b] && b - P < ANCHOR) b++; while (b > P && /\s/.test(M.raw[b - 1])) b--; }
    if (b <= a) return null;
    return { start: a, text: M.raw.slice(a, b) };
  }
  /* תווי raw של טווח תיבות נקיות [from..to] */
  function tokSpan(M, from, to) {
    const A = M.toks[from], B = M.toks[to];
    if (!A || !B) return null;
    return { a: M.rawIdx[A.cs], b: M.rawIdx[B.ce - 1] + 1 };
  }
  const escSearch = (s) => s.replace(/\^/g, '^^');

  // ───────── ה-port
  class WordPort {
    constructor(opts) {
      this.opts = opts || {};
      this.head = null; this.next = null;      // uniqueLocalId של שורת הסימן ושל הסימן הבא
      this.keys = [];                          // מפתחות פסקאות הסימן בקריאה האחרונה
      this.lastError = null; this.lastWarn = null;
      this.diag = [];
      this.me = ME;
    }
    _d(s) { this.diag.push(new Date().toLocaleTimeString('he-IL', { hour12: false }) + ' ' + s); if (this.diag.length > 80) this.diag.shift(); }
    _fail(why) { this.lastError = why; this._d('✗ ' + why); return false; }

    async getTracking() {
      return Word.run(async (ctx) => { ctx.document.load('changeTrackingMode'); await ctx.sync(); return ctx.document.changeTrackingMode !== 'Off'; });
    }
    async setTrackingOn() {
      return Word.run(async (ctx) => {
        ctx.document.load('changeTrackingMode'); await ctx.sync();
        if (ctx.document.changeTrackingMode === 'Off') { ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll; await ctx.sync(); ctx.document.load('changeTrackingMode'); await ctx.sync(); }
        return ctx.document.changeTrackingMode !== 'Off';
      });
    }

    // ───── פסקאות הסימן
    async _fullScan(ctx, name) {
      const t0 = Date.now();
      const ps = ctx.document.body.paragraphs; ps.load('items/style,items/text,items/uniqueLocalId'); await ctx.sync();
      const want = normHead('סימן ' + name);
      // משווים לטקסט הנקי (paragraph.text כולל טקסט מחוק במעקב)
      const cand = ps.items.filter((p) => p.style === SIMAN_STYLE && normHead(p.text).includes(normHead(name)));
      const rv = cand.map((p) => p.getRange('Whole').getReviewedText('Current')); await ctx.sync();
      const heads = cand.filter((p, i) => normHead(stripEnd(rv[i].value)) === want);
      if (heads.length !== 1) throw new Error(heads.length ? 'יש במסמך יותר משורת "סימן ' + name + '" אחת' : 'שורת "סימן ' + name + '" (בסגנון ' + SIMAN_STYLE + ') לא נמצאה במסמך');
      const i0 = ps.items.indexOf(heads[0]);
      let i1 = i0 + 1; while (i1 < ps.items.length && ps.items[i1].style !== SIMAN_STYLE) i1++;
      this.head = heads[0].uniqueLocalId; this.next = i1 < ps.items.length ? ps.items[i1].uniqueLocalId : null;
      this._d('סריקה מלאה: ' + ps.items.length + ' פסקאות, ' + (Date.now() - t0) + ' מ"ש');
      return ps.items.slice(i0, i1);
    }
    async _simanParas(ctx, name) {
      if (this.head) {
        try {
          const h = ctx.document.getParagraphByUniqueLocalId(this.head);
          const end = this.next ? ctx.document.getParagraphByUniqueLocalId(this.next).getRange('Start') : ctx.document.body.getRange('End');
          const ps = h.getRange('Start').expandTo(end).paragraphs; ps.load('items/style,items/text,items/uniqueLocalId');
          await ctx.sync();
          const L = ps.items;
          if (L.length && L[0].uniqueLocalId === this.head && L[0].style === SIMAN_STYLE) {
            let n = 1; while (n < L.length && L[n].style !== SIMAN_STYLE) n++;
            return L.slice(0, n);
          }
        } catch (e) { /* נופלים לסריקה מלאה */ }
      }
      return this._fullScan(ctx, name);
    }
    async getSimanParagraphs(name) {
      this.name = name;
      return Word.run(async (ctx) => {
        const L = await this._simanParas(ctx, name);
        const rv = L.map((p) => p.getRange('Whole').getReviewedText('Current'));
        await ctx.sync();
        this.keys = L.map((p) => p.uniqueLocalId);
        return L.map((p, i) => ({ key: p.uniqueLocalId, style: p.style, tokens: stripEnd(rv[i].value).split(/\s+/).filter(Boolean) }));
      });
    }
    async _para(ctx, key) {
      const L = await this._simanParas(ctx, this.name);
      const i = L.findIndex((p) => p.uniqueLocalId === key);
      if (i < 0) throw new Error('הפסקה אינה בסימן שנטען');
      return { p: L[i], index: i };
    }
    async _map(ctx, p) {
      p.load('text');
      const W = p.getRange('Whole'), rv = W.getReviewedText('Current'), tcs = W.getTrackedChanges();
      tcs.load('items/type,items/text,items/author');
      await ctx.sync();
      const raw = p.text, clean = stripEnd(rv.value);
      const dels = tcs.items.filter((t) => t.type === 'Deleted');
      let D = [];
      if (raw !== clean && dels.length) {
        const st = p.getRange('Start');
        const ms = dels.map((t) => { const r = st.expandTo(t.getRange('Start')); r.load('text'); return r; });
        await ctx.sync();
        D = dels.map((t, i) => ({ start: ms[i].text.length, len: stripEnd(t.text).length }));
      }
      return buildMap(raw, clean, D);
    }
    /* טווח Word לנקודה P ב-raw. at='end': הנקודה בסוף עוגן; 'start': בתחילתו */
    async _point(ctx, p, M, P, at) {
      const an = anchorAt(M, P, at);
      if (!an) {
        if (P === 0) return p.getRange('Start');
        if (P >= M.raw.length) return p.getRange('Content').getRange('End');
        throw new Error('אין עוגן לחיפוש');
      }
      const r = await this._find(ctx, p, an);
      return r.getRange(at === 'end' ? 'End' : 'Start');
    }
    async _find(ctx, p, an) {
      const res = p.search(escSearch(an.text), { matchCase: true, matchWildcards: false });
      res.load('items/text');
      await ctx.sync();
      if (!res.items.length) throw new Error('החיפוש לא מצא את המקום');
      const st = p.getRange('Start');
      const ms = res.items.slice(0, 300).map((r) => { const x = st.expandTo(r.getRange('Start')); x.load('text'); return x; });
      await ctx.sync();
      const hit = res.items.filter((r, i) => i < ms.length && ms[i].text.length === an.start && r.text === an.text);
      if (hit.length !== 1) throw new Error('המקום לא אותר בוודאות (' + hit.length + ')');
      return hit[0];
    }
    async _span(ctx, p, M, a, b) {
      const s = await this._point(ctx, p, M, a, 'start'), e = await this._point(ctx, p, M, b, 'end');
      const r = s.expandTo(e); r.load('text'); await ctx.sync();
      if (r.text !== M.raw.slice(a, b)) throw new Error('הטווח שאותר אינו הטקסט המצופה');
      return r;
    }
    /* לא נוגעים בהוספה במעקב של יונתן עצמו */
    async _guard(ctx, r) {
      const tcs = r.getTrackedChanges(); tcs.load('items/type,items/author'); await ctx.sync();
      const mine = tcs.items.filter((t) => t.type === 'Added' && t.author === this.me);
      if (mine.length) throw new Error('המקום כולל הוספה שלך במעקב, והתוסף לא מוחק אותה. לעשות ב-Word ידנית');
    }
    async _style(ctx, rng, runs, paraStyle) {
      if (!runs || !runs.length) return;
      for (const r of runs) {
        const t = String(r.text || '').trim(); if (!t) continue;
        if (r.char_style) {
          const res = rng.search(escSearch(t.slice(0, 250)), { matchCase: true }); res.load('items'); await ctx.sync();
          if (res.items.length) res.items[0].style = CHAR_MAP[r.char_style] || r.char_style;
          else this.lastWarn = 'סגנון התו לא הוחל על "' + t.slice(0, 20) + '": להחיל ב-Word';
        }
      }
      await ctx.sync();
    }
    async _checkInherited(ctx, rng, runs, paraStyle) {
      if (runs && runs.some((r) => r.char_style)) return;
      rng.load('style'); await ctx.sync();
      if (rng.style && paraStyle && rng.style !== paraStyle) this.lastWarn = 'הטקסט שנכתב קיבל את הסגנון "' + rng.style + '" מהטקסט שלפניו: לבדוק ב-Word';
    }

    async applyEdit(e) {
      this.lastError = null; this.lastWarn = null;
      const t0 = Date.now();
      try {
        const r = await Word.run(async (ctx) => {
          ctx.document.load('changeTrackingMode'); await ctx.sync();
          if (ctx.document.changeTrackingMode === 'Off') throw new Error('מעקב השינויים כבוי');
          const { p, index } = await this._para(ctx, e.key);
          p.load('style'); await ctx.sync();
          if (e.kind === 'insertParagraphBefore') {
            if (index === 0) throw new Error('לא מוסיפים פסקה לפני שורת הסימן');
            const np = p.insertParagraph(e.text, 'Before');
            np.style = CHAR_MAP[e.style] || e.style;
            await ctx.sync();
            await this._style(ctx, np.getRange('Content'), e.runs);
            return true;
          }
          const M = await this._map(ctx, p);
          const n = M.toks.length;
          const bad = (i) => !(i >= 0 && i < n);
          switch (e.kind) {
            case 'tokenEdit': {
              if (bad(e.token)) throw new Error('תיבה לא קיימת');
              const T = M.toks[e.token], L = T.ce - T.cs;
              if (e.start < 0 || e.start + e.len > L) throw new Error('תווים מחוץ לתיבה');
              const c0 = T.cs + e.start, c1 = c0 + e.len;
              if (e.len === 0) {
                const at = e.start === 0 ? 'start' : 'end';
                const P = e.start === 0 ? rawPoint(M, c0, 'before') : rawPoint(M, c0, 'after');
                const pt = await this._point(ctx, p, M, P, at);
                pt.insertText(e.text, 'Start'); await ctx.sync();
                return true;
              }
              const a = M.rawIdx[c0], b = M.rawIdx[c1 - 1] + 1;
              const rng = await this._span(ctx, p, M, a, b);
              await this._guard(ctx, rng);
              if (e.text) rng.insertText(e.text, 'Replace'); else rng.delete();
              await ctx.sync();
              return true;
            }
            case 'insertAfter': case 'insertBefore': {
              if (bad(e.token)) throw new Error('תיבה לא קיימת');
              // נכתב אחרי הרווח (לפני התיבה הבאה), כדי לא לרשת סגנון תו של המילה שלפני (למשל מ"מ)
              let ins;
              const nextTok = e.kind === 'insertAfter' ? e.token + 1 : e.token;
              if (nextTok < n) {
                const pt = await this._point(ctx, p, M, rawPoint(M, M.toks[nextTok].cs, 'before'), 'start');
                ins = pt.insertText(e.text + ' ', 'Start');
              } else {
                const pt = await this._point(ctx, p, M, rawPoint(M, M.toks[n - 1].ce, 'after'), 'end');
                ins = pt.insertText(' ' + e.text, 'Start');
              }
              await ctx.sync();
              await this._style(ctx, ins, e.runs);
              await this._checkInherited(ctx, ins, e.runs, p.style);
              return true;
            }
            case 'replaceTokens': case 'deleteTokens': {
              if (bad(e.from) || bad(e.to) || e.to < e.from) throw new Error('תיבות לא קיימות');
              let { a, b } = tokSpan(M, e.from, e.to);
              if (e.kind === 'deleteTokens') {     // מוחקים גם רווח אחד, כדי שלא יישאר רווח כפול
                if (e.to + 1 < n) b = M.rawIdx[M.toks[e.to + 1].cs];
                else if (e.from > 0) a = M.rawIdx[M.toks[e.from - 1].ce - 1] + 1;
              }
              const rng = await this._span(ctx, p, M, a, b);
              await this._guard(ctx, rng);
              if (e.kind === 'deleteTokens') { rng.delete(); await ctx.sync(); return true; }
              const ins = rng.insertText(e.text, 'Replace'); await ctx.sync();
              await this._style(ctx, ins, e.runs);
              await this._checkInherited(ctx, ins, e.runs, p.style);
              return true;
            }
            case 'split': {
              if (bad(e.token) || e.token === 0) throw new Error('פיצול לא אפשרי כאן');
              const T = M.toks[e.token];
              const rng = await this._span(ctx, p, M, M.rawIdx[T.cs], M.rawIdx[T.ce - 1] + 1);
              rng.insertText('\n', 'Before'); await ctx.sync();
              return true;
            }
          }
          throw new Error('פעולה לא מוכרת ' + e.kind);
        });
        this._d('✓ ' + e.kind + ' ' + (Date.now() - t0) + ' מ"ש' + (this.lastWarn ? ' · ' + this.lastWarn : ''));
        return r;
      } catch (err) {
        return this._fail((err && err.message) || String(err));
      }
    }

    // ───── בחירה, הערה, סמן
    async _selectRange(key, from, to, fn) {
      return Word.run(async (ctx) => {
        const { p } = await this._para(ctx, key);
        if (from === undefined) return fn(ctx, p.getRange('Start'));
        const M = await this._map(ctx, p);
        const s = tokSpan(M, from, Math.min(to, M.toks.length - 1));
        if (!s) return fn(ctx, p.getRange('Start'));
        return fn(ctx, await this._span(ctx, p, M, s.a, s.b));
      });
    }
    async select(key, from, to) { return this._selectRange(key, from, to, async (ctx, r) => { r.select('Select'); await ctx.sync(); }); }
    async selectParagraph(key) { return this._selectRange(key, undefined, undefined, async (ctx, r) => { r.select('Select'); await ctx.sync(); }); }
    async insertComment(key, from, to, text) {
      try { await this._selectRange(key, from, to, async (ctx, r) => { r.insertComment(text); await ctx.sync(); }); return true; }
      catch (e) { return this._fail(e.message); }
    }
    async getCursorParagraph() {
      return Word.run(async (ctx) => {
        const p = ctx.document.getSelection().paragraphs.getFirst(); p.load('uniqueLocalId'); await ctx.sync();
        return this.keys.includes(p.uniqueLocalId) ? p.uniqueLocalId : null;
      }).catch(() => null);
    }
    onCursorMoved(cb) {
      let t = null;
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => { clearTimeout(t); t = setTimeout(cb, 250); });
    }
  }

  return { WordPort, buildMap, rawPoint, anchorAt, tokSpan, CHAR_MAP };
}));
