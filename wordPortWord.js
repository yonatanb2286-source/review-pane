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
  const CHAR_MAP = { '8_מ"מ תו': '8_מ"מ', 'תרגום': '9_תרגום' };            // ממצא שלב 0: השם בפועל הוא 8_מ"מ (סגנון מקושר)
  const ANCHOR = 24;
  const stripEnd = (s) => String(s == null ? '' : s).replace(/[\r\n\u0007]+$/, '');
  const normHead = (s) => String(s).replace(/[^\p{L}\p{N}\s]/gu, '').replace(/\s+/g, ' ').trim();

  // ───────── חלק טהור (נבדק ב-Node): מיפוי בין raw ל-clean
  /* dels: [{start, len}] בקואורדינטות raw. מחזיר {live:[bool], rawIdx:[clean->raw], toks:[{cs,ce}]} או זורק אם לא מתיישב */
  /* חלופה כשמיקומי המחיקות מ-Word לא מתיישבים: יישור clean בתוך raw במינימום קטעים מחוקים.
     בשוויון מעדיפים מחיקה מוקדמת (ב-Word המחוק בא לפני המוסף). מחזיר מערך live או null. */
  function alignLive(raw, clean, cand) {
    const n = raw.length, m = clean.length, B = n - m;
    if (B < 0) return null;
    if (B === 0) return raw === clean ? new Array(n).fill(true) : null;
    const INF = 1e9, W = B + 1;
    // cost[i][d][g]: אחרי i תווי raw, d מחוקים (j=i-d), g=האם התו האחרון נמחק
    const cost = [new Float64Array(W * 2).fill(INF)];
    cost[0][0] = 0;
    for (let i = 0; i < n; i++) {
      const cur = cost[i], nx = new Float64Array(W * 2).fill(INF);
      for (let d = 0; d <= B; d++) for (let g = 0; g < 2; g++) {
        const c = cur[d * 2 + g]; if (c >= INF) continue;
        const j = i - d;
        if (d < B) { const v = c + (g ? 0 : 1) + (cand && cand[i] ? 0 : 1e-4); if (v < nx[(d + 1) * 2 + 1]) nx[(d + 1) * 2 + 1] = v; }
        if (j < m && raw[i] === clean[j] && c < nx[d * 2]) nx[d * 2] = c;
      }
      cost.push(nx);
    }
    let d = B, g = cost[n][B * 2 + 1] <= cost[n][B * 2] ? 1 : 0;
    if (cost[n][d * 2 + g] >= INF) return null;
    const live = new Array(n).fill(true);
    for (let i = n; i > 0; i--) {
      const c = cost[i][d * 2 + g];
      live[i - 1] = !g;
      if (g) { const pd = d - 1, e = (cand && cand[i - 1] ? 0 : 1e-4); const a = cost[i - 1][pd * 2 + 1], b = cost[i - 1][pd * 2]; g = Math.abs(a + e - c) < 1e-9 ? 1 : (Math.abs(b + 1 + e - c) < 1e-9 ? 0 : 1); d = pd; }
      else { const a = cost[i - 1][d * 2 + 1], b = cost[i - 1][d * 2]; g = Math.abs(b - c) < 1e-9 ? 0 : 1; if (Math.abs(a - c) > 1e-9 && Math.abs(b - c) > 1e-9) return null; }
    }
    return live;
  }
  function buildMap(raw, clean, dels, late) {
    let live = new Array(raw.length).fill(true);
    for (const d of dels) for (let k = d.start; k < d.start + d.len && k < raw.length; k++) live[k] = false;
    let rawIdx = [];
    for (let k = 0; k < raw.length; k++) if (live[k]) rawIdx.push(k);
    let how = 'tc';
    if (rawIdx.map((k) => raw[k]).join('') !== clean) {
      // רמזים: כל טקסט מחוק במקום הקרוב ביותר למיקום ש-Word דיווח
      const cand = new Array(raw.length).fill(false);
      for (const d of dels) {
        if (!d.text) continue;
        let best = -1;
        for (let k = raw.indexOf(d.text); k >= 0; k = raw.indexOf(d.text, k + 1)) if (best < 0 || Math.abs(k - d.start) < Math.abs(best - d.start)) best = k;
        if (best >= 0) for (let k = best; k < best + d.text.length; k++) cand[k] = true;
      }
      if (!late) live = alignLive(raw, clean, cand);
      else { const rv = (x) => [...x].reverse(); const l2 = alignLive(rv(raw).join(''), rv(clean).join(''), rv(cand)); live = l2 && l2.reverse(); }
      if (!live) throw new Error('טקסט הפסקה לא מתיישב עם המחיקות במעקב (raw ' + raw.length + ', נקי ' + clean.length + ')');
      rawIdx = []; for (let k = 0; k < raw.length; k++) if (live[k]) rawIdx.push(k);
      how = 'align';
    }
    const rebuilt = rawIdx.map((k) => raw[k]).join('');
    if (rebuilt !== clean) throw new Error('טקסט הפסקה לא מתיישב עם המחיקות במעקב');
    const toks = []; const re = /\S+/g; let m;
    while ((m = re.exec(clean))) toks.push({ cs: m.index, ce: m.index + m[0].length });
    return { raw, clean, live, rawIdx, toks, how };
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
    /* name: סימן אחד ("קה") או טווח ("קט-קי"): מתחילים בשורת הסימן הראשון ונגמרים לפני הסימן שאחרי האחרון */
    async _fullScan(ctx, name) {
      const t0 = Date.now();
      const names = String(name).split(/\s*[-–]\s*/).filter(Boolean);
      const ps = ctx.document.body.paragraphs; ps.load('items/style,items/text,items/uniqueLocalId'); await ctx.sync();
      const sim = ps.items.filter((p) => p.style === SIMAN_STYLE);
      const rv = sim.map((p) => p.getRange('Whole').getReviewedText('Current')); await ctx.sync();   // טקסט נקי (paragraph.text כולל מחוק)
      const find = (nm) => {
        const want = normHead('סימן ' + nm);
        const h = sim.filter((p, i) => normHead(stripEnd(rv[i].value)) === want);
        if (h.length !== 1) throw new Error(h.length ? 'יש במסמך יותר משורת "סימן ' + nm + '" אחת' : 'שורת "סימן ' + nm + '" (בסגנון ' + SIMAN_STYLE + ') לא נמצאה במסמך');
        return ps.items.indexOf(h[0]);
      };
      const i0 = find(names[0]), iLast = find(names[names.length - 1]);
      if (iLast < i0) throw new Error('סדר הסימנים במסמך הפוך');
      let i1 = iLast + 1; while (i1 < ps.items.length && ps.items[i1].style !== SIMAN_STYLE) i1++;
      this.head = ps.items[i0].uniqueLocalId; this.next = i1 < ps.items.length ? ps.items[i1].uniqueLocalId : null;
      this.nHeads = ps.items.slice(i0, i1).filter((p) => p.style === SIMAN_STYLE).length;
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
            let n = 0, k = 0;
            for (; n < L.length; n++) { if (L[n].uniqueLocalId === this.next) break; if (L[n].style === SIMAN_STYLE && ++k > (this.nHeads || 1)) break; }
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
        D = dels.map((t, i) => ({ start: ms[i].text.length, len: stripEnd(t.text).length, text: stripEnd(t.text) }));
      }
      const M = buildMap(raw, clean, D, this.lateAlign);
      if (M.how === 'align') this._d('יישור חלופי (מיקומי המחיקות מ-Word לא התיישבו): raw ' + raw.length + ', נקי ' + clean.length + ', מחיקות ' + dels.length);
      return M;
    }
    /* אחרי כתיבה של "אחר": מה שבסוגריים בטקסט שנכתב מקבל סגנון מ"מ */
    async _mm(ctx, rng) { try { await this._mm0(ctx, rng); } catch (e) { this.lastWarn = 'עיצוב הסוגריים לא הוחל: ' + e.message; } }
    async _mm0(ctx, rng) {
      if (!rng) return;
      rng.load('text'); await ctx.sync();
      // (??) שנכתב: תמיד בסגנון מ"מ, ובהזרחה צהובה כמו כותרת צד (כדי לעבור עליהם אחר כך)
      if (String(rng.text).includes('(??)')) {
        const res = rng.search('(??)', { matchCase: true, matchWildcards: false }); res.load('items'); await ctx.sync();
        for (const x of res.items) { x.style = '8_מ"מ'; if (this.highlightHeadings) x.font.highlightColor = 'Yellow'; }
        await ctx.sync();
      }
      if (!this.mmParens) return;
      const segs = String(rng.text).match(/\([^()]*\)/g) || [];
      for (const g of segs) {
        const res = rng.search(escSearch(g.slice(0, 250)), { matchCase: true }); res.load('items'); await ctx.sync();
        if (res.items.length) res.items[0].style = '8_מ"מ';
      }
      await ctx.sync();
    }
    /* טווח Word לנקודה P ב-raw. at='end': הנקודה בסוף עוגן; 'start': בתחילתו */
    async _point(ctx, p, M, P, at) {
      const an = anchorAt(M, P, at);
      if (!an) {
        if (P === 0) return p.getRange('Start');
        if (P >= M.raw.length) return p.getRange('Content').getRange('End');
        throw new Error('אין עוגן לחיפוש');
      }
      const r = await this._find(ctx, p, an, M);
      return r.getRange(at === 'end' ? 'End' : 'Start');
    }
    async _find(ctx, p, an, M) {
      const res = p.search(escSearch(an.text), { matchCase: true, matchWildcards: false });
      res.load('items/text');
      await ctx.sync();
      if (!res.items.length) throw new Error('החיפוש לא מצא את המקום');
      const st = p.getRange('Start');
      const ms = res.items.slice(0, 300).map((r) => { const x = st.expandTo(r.getRange('Start')); x.load('text'); return x; });
      await ctx.sync();
      const hit = res.items.filter((r, i) => i < ms.length && ms[i].text.length === an.start && r.text === an.text);
      if (hit.length !== 1) throw new Error('המקום לא אותר בוודאות (' + hit.length + ')');
      if (M && M.how === 'align') {     // ביישור חלופי: העוגן חייב להיות טקסט חי, בלי מחיקה במעקב בתוכו
        const tc = hit[0].getTrackedChanges(); tc.load('items/type'); await ctx.sync();
        if (tc.items.some((t) => t.type === 'Deleted')) { const er = new Error('ALIGN'); er.align = true; throw er; }
      }
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
    async _style(ctx, rng, runs, paraStyle) { try { await this._style0(ctx, rng, runs, paraStyle); } catch (e) { this.lastWarn = 'סגנון התו לא הוחל: ' + e.message; } }
    async _style0(ctx, rng, runs, paraStyle) {
      if (!runs || !runs.length) return;
      for (const r of runs) {
        const t = String(r.text || '').trim(); if (!t) continue;
        if (r.char_style) {
          const name = CHAR_MAP[r.char_style] || r.char_style;
          if (name === '8_מ"מ' && !/[()[\]]/.test(t)) continue;   // מ"מ רק למראה מקום בסוגריים, לא לשבר מילה
          const res = rng.search(escSearch(t.slice(0, 250)), { matchCase: true }); res.load('items'); await ctx.sync();
          if (!res.items.length) { this.lastWarn = 'סגנון התו לא הוחל על "' + t.slice(0, 20) + '": להחיל ב-Word'; continue; }
          // סגנון שלא קיים במסמך לא מכשיל את הכתיבה (הטקסט כבר נכתב): מנסים גם את שם סגנון התו המקושר, ואחרת מזהירים
          let ok = false;
          for (const nm of [name, name + ' תו', name + ' Char']) {
            try { res.items[0].style = nm; await ctx.sync(); ok = true; break; } catch (e) { /* הבא */ }
          }
          if (!ok) this.lastWarn = 'הסגנון "' + name + '" לא נמצא במסמך, והטקסט "' + t.slice(0, 20) + '" נכתב בלי סגנון: להחיל ב-Word';
        }
      }
    }
    async _checkInherited(ctx, rng, runs, paraStyle) { try { await this._checkInh0(ctx, rng, runs, paraStyle); } catch (e) { /* בדיקה בלבד */ } }
    async _checkInh0(ctx, rng, runs, paraStyle) {
      if (runs && runs.some((r) => r.char_style)) return;
      rng.load('style'); await ctx.sync();
      if (rng.style && paraStyle && rng.style !== paraStyle) this.lastWarn = 'הטקסט שנכתב קיבל את הסגנון "' + rng.style + '" מהטקסט שלפניו: לבדוק ב-Word';
    }

    async applyEdit(e) {
      this.lastError = null; this.lastWarn = null;
      const t0 = Date.now();
      try {
        let r;
        for (const late of [false, true]) {
          this.lateAlign = late;
          try { r = await this._apply(e); break; }
          catch (er) { if (!er.align || late) throw (er.align ? new Error('המקום לא אותר בוודאות (יישור)') : er); this._d('יישור חלופי: ניסיון שני'); }
        }
        this.lateAlign = false;
        this._d('✓ ' + e.kind + ' ' + (Date.now() - t0) + ' מ"ש' + (this.lastWarn ? ' · ' + this.lastWarn : ''));
        return r;
      } catch (err) {
        this.lateAlign = false;
        return this._fail((err && err.message) || String(err));
      }
    }
    async _apply(e) {
      {
        const r = await Word.run(async (ctx) => {
          ctx.document.load('changeTrackingMode'); await ctx.sync();
          if (ctx.document.changeTrackingMode === 'Off') throw new Error('מעקב השינויים כבוי');
          const { p, index } = await this._para(ctx, e.key);
          p.load('style'); await ctx.sync();
          if (e.kind === 'insertParagraphBefore') {
            if (index === 0) throw new Error('לא מוסיפים פסקה לפני שורת הסימן');
            const np = p.insertParagraph(e.text, 'Before');
            await ctx.sync();
            try { np.style = CHAR_MAP[e.style] || e.style; await ctx.sync(); } catch (er) { this.lastWarn = 'סגנון הפסקה "' + e.style + '" לא נמצא במסמך: להחיל ב-Word'; }
            if (this.highlightHeadings && e.style === '7_כותרת-צד') { np.getRange('Content').font.highlightColor = 'Yellow'; await ctx.sync(); }
            await this._style(ctx, np.getRange('Content'), e.runs);
            await this._mm(ctx, np.getRange('Content'));
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
                const ins0 = pt.insertText(e.text, 'Start'); await ctx.sync();
                await this._style(ctx, ins0, e.runs);
                await this._mm(ctx, ins0);
                return true;
              }
              const a = M.rawIdx[c0], b = M.rawIdx[c1 - 1] + 1;
              const rng = await this._span(ctx, p, M, a, b);
              await this._guard(ctx, rng);
              const ins1 = e.text ? rng.insertText(e.text, 'Replace') : (rng.delete(), null);
              await ctx.sync();
              if (ins1) await this._style(ctx, ins1, e.runs);
              await this._mm(ctx, ins1);
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
              await this._mm(ctx, ins);
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
              await this._mm(ctx, ins);
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
        return r;
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
    /* סימון המקום ב-Word בקו תחתון של "הערת עריכה" (annotation, WordApi 1.7): לא טקסט ולא עיצוב, ולא נכתב כשינוי.
       key=null: רק מנקה את הסימון הקודם. */
    async mark(key, from, to) {
      if (typeof Office !== 'undefined' && !Office.context.requirements.isSetSupported('WordApi', '1.7')) return;
      return Word.run(async (ctx) => {
        if (this.markIds && this.markIds.length) {
          for (const id of this.markIds) { try { ctx.document.getAnnotationById(id).delete(); } catch (e) { /* */ } }
          this.markIds = []; await ctx.sync().catch(() => {});
        }
        if (!key) return;
        const { p } = await this._para(ctx, key);
        const M = await this._map(ctx, p);
        if (!M.toks.length) return;
        const f = Math.min(from === undefined ? 0 : from, M.toks.length - 1), t = Math.min(to === undefined ? f : to, M.toks.length - 1);
        // מיקומי ה-annotation נספרים בטקסט הנקי (בלי מחוק במעקב) או ב-raw, לפי הגרסה: מנסים את שניהם ובודקים מה סומן בפועל
        const cs = M.toks[f].cs, ce = M.toks[t].ce, want = M.clean.slice(cs, ce);
        const sp = tokSpan(M, f, t);
        const tries = [[cs, ce - cs]]; if (sp && sp.a !== cs) tries.push([sp.a, sp.b - sp.a]);
        for (const [a, len] of tries) {
          const r = p.insertAnnotations({ critiques: [{ colorScheme: 'Lavender', start: a, length: Math.max(1, len) }] });
          await ctx.sync();
          const ids = r.value || [];
          let ok = false;
          try {
            const rg = ids.length ? ctx.document.getAnnotationById(ids[0]).critiqueAnnotation.range : null;
            if (rg) { rg.load('text'); await ctx.sync(); ok = rg.text.replace(/\s+/g, ' ').trim() === want.replace(/\s+/g, ' ').trim(); }
          } catch (e) { ok = false; }
          if (ok) { this.markIds = ids; return; }
          for (const id of ids) { try { ctx.document.getAnnotationById(id).delete(); } catch (e) { /* */ } }
          await ctx.sync().catch(() => {});
        }
        this._d('סימון ב-Word: המקום לא אומת, לא סומן');
      }).catch((e) => this._d('סימון ב-Word נכשל: ' + e.message));
    }
    /* ביטול אישור מהחלונית: דוחה ב-Word רק שינויים במעקב של יונתן עצמו, בתיבות של הפריט (ותיבה אחת מכל צד),
       שנוצרו בחלון הזמן של האישור. שינוי שעשה ידנית בזמן אחר לא נוגעים בו. */
    async undoOwn(key, from, to, t0, t1) {
      this.lastError = null;
      try {
        return await Word.run(async (ctx) => {
          const { p } = await this._para(ctx, key);
          const M = await this._map(ctx, p), n = M.toks.length;
          if (!n) throw new Error('הפסקה ריקה');
          const f = Math.max(0, Math.min(from, n - 1) - 1), t = Math.min(n - 1, Math.max(to, from) + 1);
          const sp = tokSpan(M, f, t);
          // כולל טקסט מחוק במעקב שצמוד לטווח (בהחלפה Word שם את המחוק ליד החדש)
          const live = new Set(M.rawIdx); let a = sp.a, b = sp.b;
          while (a > 0 && !live.has(a - 1)) a--;
          while (b < M.raw.length && !live.has(b)) b++;
          let rng;
          try { rng = await this._span(ctx, p, M, a, b); } catch (e) { rng = await this._span(ctx, p, M, sp.a, sp.b); }
          const tcs = rng.getTrackedChanges(); tcs.load('items/type,items/author,items/date'); await ctx.sync();
          const mine = tcs.items.filter((x) => x.author === this.me && x.date && +new Date(x.date) >= t0 && +new Date(x.date) <= t1);
          if (!mine.length) throw new Error('לא נמצא במקום שינוי שהחלונית כתבה בזמן האישור');
          mine.forEach((x) => x.reject()); await ctx.sync();
          this._d('↶ בוטלו ' + mine.length + ' שינויים');
          return mine.length;
        });
      } catch (e) { return this._fail(e.message); }
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

  return { WordPort, buildMap, alignLive, rawPoint, anchorAt, tokSpan, CHAR_MAP };
}));
