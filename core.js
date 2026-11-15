/* core.js — ליבת הלוגיקה של חלונית הסקירה (שלב 2). JavaScript טהור, בלי Office.js.
   מדבר עם המסמך רק דרך port (wordPort האמיתי, או wordPortFake בבדיקות ובמצב הדגמה).

   חוזה ה-port (כל הפונקציות אסינכרוניות):
     getTracking() / setTrackingOn()
     getSimanParagraphs(simanName) -> [{key, style, tokens:[טקסט נקי של כל תיבה, בלי מחיקות במעקב]}]
     applyEdit(edit) -> true/false. edit.kind אחד מ:
       tokenEdit {key, token, start, len, text}      שינוי תווים בתוך תיבה אחת (פסיק בסוף = start=len(token), len=0)
       insertAfter {key, token, text, runs}           הוספת תיבות אחרי תיבה (עם רווח)
       insertBefore {key, token, text, runs}          הוספת תיבות לפני תיבה
       replaceTokens {key, from, to, text, runs}      החלפת תיבות שלמות (כולל)
       deleteTokens {key, from, to}                   מחיקת תיבות שלמות
       split {key, token}                             פסקה חדשה מתיבה זו (באותו סגנון)
       insertParagraphBefore {key, text, style, runs} פסקה חדשה לפני הפסקה
     לכל edit יש גם tag (מזהה האשכול או הפריט), לבדיקות בלבד.
     insertComment(key, from, to, text), select(key, from, to), selectParagraph(key)

   מצב פריט נגזר מהמסמך: open / accepted / other / changed / notfound, ומהיומן: rejected / seen.
   כלל: לא כותבים אם המקום לא נמצא בוודאות, או אם האשכול "השתנה בקובץ". */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReviewCore = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const HEAD_STYLES = new Set(['7_כותרת-צד', '7-']);
  const TOPIC_STYLES = new Set(['1.1_נושא הסימן', '11']);

  // ───────── נרמול
  const canon = (t) => String(t).replace(/[֑-ׇ]/g, '').replace(/[״“”„]/g, '"').replace(/[׳‘’`]/g, "'");
  const norm = (t) => canon(t).replace(/[^\p{L}\p{N}]/gu, '') || canon(t);   // תיבת פיסוק בלבד (למשל ??) נשארת כמות שהיא
  const splitWords = (s) => String(s).split(/\s+/).filter(Boolean);

  // ───────── החלת פריטים (מקביל ל-make_review_pkg.apply_edits)
  function applyCharEdits(text, edits) {
    const e = edits.slice().sort((a, b) => (b.start - a.start) || (b.seq - a.seq));
    for (const x of e) text = text.slice(0, x.start) + x.text + text.slice(x.start + x.len);
    return text;
  }

  // ───────── טעינת חבילה
  function loadPackage(pkg) {
    if (!pkg || pkg.format !== 1) throw new Error('חבילה לא מוכרת (format)');
    const items = new Map(pkg.items.map((it) => [it.id, it]));
    const clusters = new Map(pkg.clusters.map((c) => [c.cid, c]));
    const paras = new Map(pkg.paragraphs.map((p) => [p.pid, p]));
    const clustersOf = new Map();
    for (const c of pkg.clusters) {
      if (!clustersOf.has(c.pid)) clustersOf.set(c.pid, []);
      clustersOf.get(c.pid).push(c);
      c.n = c.items.length;
      c.outcomes = [];             // לפי מסכה: מערך תיבות
      for (let m = 0; m < (1 << c.n); m++) c.outcomes.push(splitWords(outcomeText(c, items, m)));
      c.outKeys = c.outcomes.map((o) => o.map(canon).join(' '));
    }
    for (const list of clustersOf.values()) list.sort((a, b) => a.from - b.from);
    for (const p of pkg.paragraphs) {
      p.normTokens = p.orig_tokens.map(norm).filter(Boolean);
      p.normProp = (p.proposed || []).filter((x) => x.style === p.style).flatMap((x) => splitWords(x.text)).map(norm).filter(Boolean);
      p.isHeading = HEAD_STYLES.has(p.style) || HEAD_STYLES.has(p.style_id);
      p.nSplits = p.items.filter((i) => items.get(i).op === 'split_before').length;
    }
    return { pkg, items, clusters, paras, clustersOf };
  }

  function outcomeText(c, items, mask, custom) {
    // custom: {itemId: טקסט חלופי} ל"אחר"
    const eds = [];
    c.items.forEach((iid, k) => {
      if (!(mask >> k & 1)) return;
      const it = items.get(iid);
      eds.push({ start: it.rel_start, len: it.char_len, text: custom && iid in custom ? custom[iid] : it.new, seq: k });
    });
    return applyCharEdits(c.orig, eds);
  }

  // צפוי: טקסט הפסקאות אחרי אישור קבוצת פריטים (לבדיקות, מקביל ל-apply_package בפייתון)
  function expectedParagraphs(P, pid, accepted) {
    const par = P.paras.get(pid);
    const seq = [];
    const cls = P.clustersOf.get(pid) || [];
    let t = 0, ci = 0;
    while (t < par.orig_tokens.length) {
      if (ci < cls.length && cls[ci].from === t) {
        const c = cls[ci];
        let m = 0; c.items.forEach((iid, k) => { if (accepted.has(iid)) m |= 1 << k; });
        c.outcomes[m].forEach((w, j) => seq.push([j === 0 ? c.from : null, w]));
        t = c.to + 1; ci++;
      } else { seq.push([t, par.orig_tokens[t]]); t++; }
    }
    const its = par.items.map((i) => P.items.get(i));
    const splits = new Set(its.filter((x) => x.op === 'split_before' && accepted.has(x.id)).map((x) => x.from));
    const out = [{ style: par.style, tokens: [], start: 0 }];
    for (const [oi, w] of seq) {
      if (oi !== null && splits.has(oi) && out[out.length - 1].tokens.length) out.push({ style: par.style, tokens: [], start: oi });
      out[out.length - 1].tokens.push(w);
    }
    const ins = its.filter((x) => ['heading_before', 'topic', 'para_before'].includes(x.op) && accepted.has(x.id));
    const res = [];
    for (const o of out) {
      ins.filter((x) => x.from === o.start).sort((a, b) => (a.order || 0) - (b.order || 0))
        .forEach((x) => res.push({ style: x.para_style, tokens: splitWords(x.new) }));
      res.push({ style: o.style, tokens: o.tokens });
    }
    return res;
  }

  // ───────── התאמת פסקאות
  function bagSim(a, b) {
    if (!a.length && !b.length) return 1;
    const m = new Map();
    for (const x of a) m.set(x, (m.get(x) || 0) + 1);
    let common = 0;
    for (const x of b) { const v = m.get(x); if (v) { common++; m.set(x, v - 1); } }
    return 2 * common / (a.length + b.length);
  }

  /* live: [{key, style, tokens}] בסדר המסמך. מחזיר לכל pid: {keys, extras, ratio, status} */
  function matchParagraphs(P, live) {
    const L = live.map((p) => ({ ...p, norm: p.tokens.map(norm).filter(Boolean),
      isHeading: HEAD_STYLES.has(p.style) || TOPIC_STYLES.has(p.style) }));
    const res = new Map();
    let j = 0;
    for (const p of P.pkg.paragraphs) {
      let best = null;
      const maxK = 1 + p.nSplits * 2 + 3;
      const cnt0 = new Map(), cnt1 = new Map();
      for (const x of p.normTokens) cnt0.set(x, (cnt0.get(x) || 0) + 1);
      for (const x of p.normProp) cnt1.set(x, (cnt1.get(x) || 0) + 1);
      for (let s = j; s < Math.min(L.length, j + 6); s++) {
        const r0 = new Map(cnt0), r1 = new Map(cnt1);
        let c0 = 0, c1 = 0, n = 0, keys = [], inner = [];
        for (let k = 0, body = 0; body < maxK && k < maxK + 40 && s + k < L.length; k++) {
          const lp = L[s + k];
          if (!p.isHeading && lp.isHeading && k > 0) { inner.push(lp.key); continue; }   // כותרות שנוספו לא נספרות
          body++;
          if (k === 0 && !p.isHeading && lp.isHeading) break;
          for (const x of lp.norm) {
            const v0 = r0.get(x); if (v0) { c0++; r0.set(x, v0 - 1); }
            const v1 = r1.get(x); if (v1) { c1++; r1.set(x, v1 - 1); }
          }
          n += lp.norm.length; keys = keys.concat([lp.key]);
          const sim = (c, m) => (m + n) ? 2 * c / (m + n) : 1;
          const r = Math.max(sim(c0, p.normTokens.length), p.normProp.length ? sim(c1, p.normProp.length) : 0);
          const cand = { s, e: s + k + 1, keys: keys.slice(), inner: inner.slice(), ratio: r };
          if (!best || r > best.ratio + 1e-9 || (Math.abs(r - best.ratio) < 1e-9 && s < best.s)) best = cand;
          if (p.isHeading) break;
        }
      }
      if (best && best.ratio >= 0.6) {
        // פסקאות שלפני ההתאמה (כותרות שנוספו, פסקאות של יונתן)
        const extras = L.slice(j, best.s).map((x) => x.key);
        res.set(p.pid, { keys: best.keys, inner: best.inner, extras, ratio: best.ratio,
          status: best.ratio >= 0.9 ? 'ok' : 'changed', first: best.s, last: best.e - 1 });
        j = best.e;
      } else {
        res.set(p.pid, { keys: [], inner: [], extras: [], ratio: best ? best.ratio : 0, status: 'missing' });
      }
    }
    return { map: res, live: L };
  }

  // ───────── איתור אשכול ברצף התיבות החי של הפסקה
  function concatTokens(match, liveByKey) {
    const seq = [];       // {key, idx, t, n}
    for (const key of match.keys) {
      const lp = liveByKey.get(key);
      lp.tokens.forEach((t, idx) => seq.push({ key, idx, t, n: norm(t) }));
    }
    return seq;
  }

  /* outcomes: מערכי תיבות אפשריים. מחזיר {pos, len, mask} או null. pos/len בתוך seq. */
  function locate(seq, ctxB, ctxA, outcomes, expectedFrac) {
    const nz = [];                     // אינדקסים בעלי ערך מנורמל
    seq.forEach((x, i) => { if (x.n) nz.push(i); });
    const nzVals = nz.map((i) => seq[i].n);
    const cb = ctxB.map(norm).filter(Boolean), ca = ctxA.map(norm).filter(Boolean);
    const cands = [];
    const seen = new Set();
    outcomes.forEach((o, mask) => {
      const on = o.map(norm).filter(Boolean);
      const key = on.join(' ');
      if (seen.has(key)) return; seen.add(key);
      for (let s = 0; s + on.length <= nzVals.length; s++) {
        let ok = true;
        for (let k = 0; k < on.length; k++) if (nzVals[s + k] !== on[k]) { ok = false; break; }
        if (!ok) continue;
        // הקשר: רצוף (חזק) ובחלון של שתי תיבות נוספות (עמיד לשינוי מילה סמוכה, למשל פתיחת ר"ת)
        let sb = 0; for (let k = 1; k <= cb.length && s - k >= 0; k++) { if (nzVals[s - k] === cb[cb.length - k]) sb++; else break; }
        let sa = 0; for (let k = 0; k < ca.length && s + on.length + k < nzVals.length; k++) { if (nzVals[s + on.length + k] === ca[k]) sa++; else break; }
        const wb = nzVals.slice(Math.max(0, s - cb.length - 1), s), wa = nzVals.slice(s + on.length, s + on.length + ca.length + 1);
        const fb = cb.filter((x) => wb.includes(x)).length, fa = ca.filter((x) => wa.includes(x)).length;
        // גבולות ב-seq: מהתיבה המנורמלת הראשונה; תיבות פיסוק בלבד שבתוך הטווח נכללות
        const startSeq = on.length ? nz[s] : (s < nz.length ? nz[s] : seq.length);
        const endSeq = on.length ? nz[s + on.length - 1] + 1 : startSeq;
        const exact = seq.slice(startSeq, endSeq).map((x) => canon(x.t)).join(' ') === o.map(canon).join(' ');
        cands.push({ s, startSeq, endSeq, score: fb + fa, strict: sb + sa, sb, sa, on, exact, empty: !on.length });
      }
    });
    if (!cands.length) return null;
    // גבול פסקה נחשב הקשר תואם כשההקשר המקורי קצר בגללו
    for (const c of cands) {
      c.bonus = (cb.length < 3 && c.s === cb.length && c.sb === cb.length ? 1 : 0)
              + (ca.length < 3 && c.s + c.on.length === nzVals.length - ca.length && c.sa === ca.length ? 1 : 0);
      c.score += c.bonus; c.strict += c.bonus;
    }
    const need = Math.min(2, cb.length + ca.length + (cb.length < 3 ? 1 : 0) + (ca.length < 3 ? 1 : 0));
    // לכל מקום: המועמד הטוב (נוסח מדויק, ארוך)
    const byPos = new Map();
    for (const c of cands) {
      const u = byPos.get(c.startSeq);
      if (!u || c.score > u.score || (c.score === u.score && (c.strict > u.strict || (c.strict === u.strict && (c.exact > u.exact || (c.exact === u.exact && c.endSeq > u.endSeq)))))) byPos.set(c.startSeq, c);
    }
    const all = [...byPos.values()].sort((a, b) => (b.score - a.score) || (b.strict - a.strict));
    const best = all[0];
    if (best.score < need) return null;                       // אין די הקשר: לא כותבים
    if (best.strict < 1 && cb.length + ca.length > 0) return null;   // חייבת להיות תיבה צמודה תואמת לפחות מצד אחד
    const rivals = all.slice(1).filter((c) => c.score >= best.score - 1);
    if (rivals.length && !(best.strict >= 2 && rivals.every((c) => c.strict <= best.strict - 2))) return null;   // דו-משמעי
    void expectedFrac;
    return best;
  }

  /* איתור לפי ההקשר בלבד, כשהטקסט במקום אינו אף אחת מהאפשרויות (עריכה של יונתן, "אחר").
     דורש את כל ההקשר משני הצדדים (או תחילת/סוף פסקה), רצוף, ומועמד יחיד. */
  function locateBetween(seq, ctxB, ctxA, maxLen) {
    const nz = []; seq.forEach((x, i) => { if (x.n) nz.push(i); });
    const v = nz.map((i) => seq[i].n);
    const cb = ctxB.map(norm).filter(Boolean), ca = ctxA.map(norm).filter(Boolean);
    if (cb.length + ca.length < 2) return null;
    const startsB = [];
    if (!cb.length) startsB.push(0);
    else for (let s = cb.length; s <= v.length; s++) if (cb.every((x, k) => v[s - cb.length + k] === x)) startsB.push(s);
    const cands = [];
    for (const s of startsB) {
      for (let L = 1; L <= maxLen && s + L <= v.length; L++) {
        const e = s + L;
        if (ca.length ? ca.every((x, k) => v[e + k] === x) : e === v.length) { cands.push({ s, e }); break; }
      }
    }
    if (cands.length !== 1) return null;
    const c = cands[0];
    return { startSeq: nz[c.s], endSeq: nz[c.e - 1] + 1, exact: false, between: true };
  }

  // ───────── מצב
  class State {
    constructor(P, log) { this.P = P; this.log = log || {}; }

    /* מחשב מצב לכל פריט מול המסמך החי. מחזיר Map id -> {state, loc, ...} */
    compute(live, onlyPids) {
      const P = this.P;
      const M = matchParagraphs(P, live);
      const liveByKey = new Map(M.live.map((p) => [p.key, p]));
      const keyIndex = new Map(M.live.map((p, i) => [p.key, i]));
      const st = onlyPids && this.st ? this.st : new Map();
      this.match = M; this.liveByKey = liveByKey; this.keyIndex = keyIndex;
      for (const p of P.pkg.paragraphs) {
        if (onlyPids && this.st && !onlyPids.has(p.pid)) continue;
        const m = M.map.get(p.pid);
        const its = p.items.map((i) => P.items.get(i));
        if (m.status === 'missing') { its.forEach((it) => st.set(it.id, { state: 'notfound', why: 'הפסקה לא נמצאה' })); continue; }
        const seq = concatTokens(m, liveByKey);
        const ntok = Math.max(1, p.orig_tokens.length);
        const clLoc = new Map();
        /* מיקום תחילת התיבה המקורית t ב-seq: דרך האשכול שמתחיל בה, או התיבה עצמה (שלא השתנתה) עם ההקשר */
        const tokenStart = (t) => {
          const c = (P.clustersOf.get(p.pid) || []).find((x) => x.from <= t && t <= x.to);
          if (c) { const l = clLoc.get(c.cid); return (c.from === t && l) ? l.startSeq : null; }
          const l = locate(seq, p.orig_tokens.slice(Math.max(0, t - 3), t), p.orig_tokens.slice(t + 1, t + 4), [[p.orig_tokens[t]]], t / ntok);
          return l && l.exact ? l.startSeq : null;
        };
        for (const c of (P.clustersOf.get(p.pid) || [])) {
          // נוסח "אחר" שנרשם ביומן הוא אפשרות נוספת במקום (אחרת תוצאה קצרה יותר, למשל המקור, עלולה להיתפס כקידומת שלו)
          const extra = c.items.map((iid) => this.log[iid]).filter((l) => l && l.state === 'other' && l.span).map((l) => splitWords(l.span));
          let loc = locate(seq, c.ctx_before, c.ctx_after, c.outcomes.concat(extra), c.from / ntok);
          if (!loc) loc = locateBetween(seq, c.ctx_before, c.ctx_after, c.outcomes[0].length + 6);   // נוסח אחר במקום (של יונתן)
          clLoc.set(c.cid, loc);
          if (!loc) { c.items.forEach((iid) => st.set(iid, { state: 'notfound', why: 'המקום לא נמצא בוודאות' })); continue; }
          // הטווח המדויק: כל התיבות בין startSeq ל-endSeq, באותה פסקה חיה
          const span = seq.slice(loc.startSeq, loc.endSeq);
          const sameKey = span.every((x) => x.key === (span[0] || {}).key);
          const exact = span.map((x) => canon(x.t)).join(' ');
          const mask = c.outKeys.indexOf(exact);
          const where = span.length ? { key: span[0].key, from: span[0].idx, to: span[span.length - 1].idx }
            : (loc.startSeq < seq.length ? { key: seq[loc.startSeq].key, from: seq[loc.startSeq].idx, to: seq[loc.startSeq].idx - 1 } : null);
          c.items.forEach((iid, k) => {
            const lg = this.log[iid];
            if (!sameKey || !where) { st.set(iid, { state: 'notfound', why: 'האשכול חוצה פסקאות' }); return; }
            if (mask >= 0) {
              const on = !!(mask >> k & 1);
              let s = on ? 'accepted' : 'open';
              if (!on && lg && lg.state === 'rejected') s = 'rejected';
              if (!on && lg && lg.state === 'other' && lg.wide && lg.span === exact) s = 'other';   // "אחר" רחב: יונתן שינה סביב ההצעה (למשל הזיז פסיק)
              if (on && lg && lg.state === 'other') s = 'other';   // "אחר" שהנוסח שלו זהה להצעה
              st.set(iid, { state: s, loc: where, mask, cid: c.cid, current: exact });
            } else {
              // השתנה בקובץ: אם ביומן "אחר" עם הנוסח הזה – אחר
              const otherOk = lg && lg.state === 'other' && lg.span === exact;
              st.set(iid, { state: otherOk ? 'other' : 'changed', loc: where, cid: c.cid, current: exact });
            }
          });
        }
        // פיצולים, כותרות, נושא, מידע
        for (const it of its) {
          if (it.cluster) continue;
          if (it.op === 'none') {
            // פריט מידע: גם לו יש מקום, כדי שיונתן יוכל לכתוב שם "אחר"
            const lg = this.log[it.id], outs = [p.orig_tokens.slice(it.from, it.to + 1)];
            if (lg && lg.state === 'other' && lg.span) outs.push(splitWords(lg.span));
            const l = locate(seq, p.orig_tokens.slice(Math.max(0, it.from - 3), it.from), p.orig_tokens.slice(it.to + 1, it.to + 4), outs, it.from / ntok);
            let loc = null, current = null;
            if (l) { const sp = seq.slice(l.startSeq, l.endSeq); if (sp.length && sp.every((x) => x.key === sp[0].key)) { loc = { key: sp[0].key, from: sp[0].idx, to: sp[sp.length - 1].idx }; current = sp.map((x) => canon(x.t)).join(' '); } }
            st.set(it.id, { state: (lg && lg.state) || 'info', loc, current });
            continue;
          }
          if (it.op === 'split_before') {
            const ss = tokenStart(it.from);
            if (ss === null || ss >= seq.length) { st.set(it.id, { state: 'notfound', why: 'מקום הפיצול לא נמצא' }); continue; }
            const x = seq[ss];
            const acc = x.idx === 0 && x.key !== m.keys[0];
            st.set(it.id, { state: acc ? 'accepted' : (this.log[it.id] && this.log[it.id].state === 'rejected' ? 'rejected' : 'open'),
              loc: { key: x.key, token: x.idx } });
            continue;
          }
          // פסקה שלמה לפני תיבה (כותרת צד / נושא)
          let anchorKey = m.keys[0], anchorOk = true;
          {   // גם לכותרת בתחילת הפסקה: התיבה הראשונה המקורית חייבת להימצא ולהיות ראשונה בפסקה החיה
            const ss = tokenStart(it.from);
            if (ss === null || ss >= seq.length) anchorOk = false;
            else { const x = seq[ss]; anchorKey = x.key; if (x.idx !== 0) anchorOk = (it.from === 0 || !(it.depends_on || []).length) ? false : 'nosplit'; }
          }
          if (anchorOk === false) { st.set(it.id, { state: 'notfound', why: 'מקום הכותרת לא נמצא' }); continue; }
          const ai = keyIndex.get(anchorKey);
          const prev = ai > 0 ? M.live[ai - 1] : null;
          const wantStyle = new Set([it.para_style, it.para_style_id]);
          const isNew = prev && wantStyle.has(prev.style) && !this._isPkgPara(prev.key);
          const lg = this.log[it.id];
          let s;
          if (anchorOk === 'nosplit') s = (lg && lg.state === 'rejected') ? 'rejected' : 'blocked';
          else if (isNew && norm(prev.tokens.join(' ')) === norm(it.new)) s = (lg && lg.state === 'other') ? 'other' : 'accepted';
          else if (isNew) s = (lg && lg.state === 'other' && norm(lg.span || '') === norm(prev.tokens.join(' '))) ? 'other' : 'changed';
          else s = (lg && lg.state === 'rejected') ? 'rejected' : 'open';
          st.set(it.id, { state: s, loc: { key: anchorKey }, current: isNew ? prev.tokens.join(' ') : null });
        }
      }
      // תלויות: פריט שתלוי בפריט שנדחה – לא רלוונטי; שתלוי בפריט פתוח – חסום
      for (const it of P.pkg.items) {
        const s = st.get(it.id); if (!s) continue;
        for (const d of it.depends_on || []) {
          const ds = st.get(d);
          if (!ds) continue;
          if (ds.state === 'rejected' && (s.state === 'open' || s.state === 'blocked')) s.state = 'na';
          else if (!['accepted', 'other'].includes(ds.state) && s.state === 'open') s.state = 'blocked';
        }
        if (s.state === 'blocked' && !(it.depends_on || []).length) s.state = 'open';
      }
      this.st = st;
      return st;
    }
    _isPkgPara(key) {
      for (const v of this.match.map.values()) if (v.keys.includes(key)) return true;
      return false;
    }
  }

  // ───────── תכנון כתיבה מזערית: מתיבות נוכחיות לתיבות יעד
  function planTokenEdit(key, from, cur, tgt) {
    const sc = cur.join(' '), stg = tgt.join(' ');
    if (sc === stg) return null;
    let pre = 0; while (pre < sc.length && pre < stg.length && sc[pre] === stg[pre]) pre++;
    let suf = 0; while (suf < sc.length - pre && suf < stg.length - pre && sc[sc.length - 1 - suf] === stg[stg.length - 1 - suf]) suf++;
    const oA = pre, oB = sc.length - suf;           // [oA,oB) בנוכחי
    const nm = stg.slice(pre, stg.length - suf);
    // מיקומי התיבות בנוכחי
    const pos = []; let x = 0; for (const t of cur) { pos.push([x, x + t.length]); x += t.length + 1; }
    const tokAt = (c) => pos.findIndex(([a, b]) => a <= c && c <= b);
    const om = sc.slice(oA, oB);
    // (א) בתוך תיבה אחת
    if (!/\s/.test(om) && cur.length && (!/\s/.test(nm) || (om === '' && !nm.startsWith(' ') && !nm.endsWith(' ')) || (om === '' && pos.some(([a, b]) => a < oA && oA < b)))) {
      let t = tokAt(oA);
      if (t >= 0 && oB <= pos[t][1]) {
        return { kind: 'tokenEdit', key, token: from + t, start: oA - pos[t][0], len: oB - oA, text: nm };
      }
    }
    // (ב) הוספת תיבות שלמות
    if (om === '' && nm.length) {
      if (nm.startsWith(' ')) { const t = pos.findIndex(([a, b]) => b === oA); if (t >= 0) return { kind: 'insertAfter', key, token: from + t, text: nm.trim() }; }
      if (nm.endsWith(' ')) { const t = pos.findIndex(([a]) => a === oA); if (t >= 0) return { kind: 'insertBefore', key, token: from + t, text: nm.trim() }; }
      if (!cur.length) return null;
    }
    // (ג) הרחבה לתיבות שלמות: החלפה או מחיקה
    if (!cur.length) return null;
    let t0 = pos.findIndex(([a, b]) => b >= oA); if (t0 < 0) t0 = pos.length - 1;
    let t1 = t0; for (let i = pos.length - 1; i >= t0; i--) if (pos[i][0] < Math.max(oB, oA + 1)) { t1 = i; break; }
    while (t0 > 0 && pos[t0][0] > oA) t0--;
    while (t1 < pos.length - 1 && pos[t1][1] < oB) t1++;
    const A = pos[t0][0], B = pos[t1][1];
    const newText = stg.slice(A, stg.length - (sc.length - B));
    if (!newText.trim()) return { kind: 'deleteTokens', key, from: from + t0, to: from + t1 };
    return { kind: 'replaceTokens', key, from: from + t0, to: from + t1, text: newText.trim() };
  }

  // ───────── הסשן: חבילה + port + יומן
  class Session {
    constructor(port, pkg, opts) {
      this.port = port; this.P = loadPackage(pkg); this.opts = opts || {};
      this.log = (opts && opts.log) || {};          // id -> {state, time, text, why, span}
      this.S = new State(this.P, this.log);
      this.busy = false;
    }
    /* pid: לחשב מחדש רק פסקה אחת (אחרי כתיבה בה). בלי pid: הכול */
    async refresh(pid) {
      this.live = await this.port.getSimanParagraphs(this.P.pkg.siman);
      return this.S.compute(this.live, pid ? new Set([pid]) : null);
    }
    state(id) { return this.S.st.get(id); }

    async _write(edit) {
      if (this.busy) throw new Error('כתיבה אחרת בעבודה');
      this.busy = true;
      try {
        if (!(await this.port.getTracking())) { await this.port.setTrackingOn(); if (!(await this.port.getTracking())) return { ok: false, why: 'מעקב השינויים כבוי ולא נדלק: לא נכתב דבר' }; }
        const ok = await this.port.applyEdit(edit);
        return { ok: ok !== false };
      } finally { this.busy = false; }
    }

    /* אישור פריט (או "אחר" עם custom: טקסט חלופי לפריט). מחזיר {ok, why, edit} */
    async approve(id, custom) {
      const P = this.P, it = P.items.get(id);
      if (!it) return { ok: false, why: 'אין פריט כזה' };
      await this.refresh(this.S.st ? it.pid : null);
      const s = this.state(id);
      if (it.op === 'none') {
        if (custom === undefined) return { ok: false, why: 'פריט מידע: אין מה לכתוב' };
        if (!s || !s.loc) return { ok: false, why: 'המקום לא נמצא בוודאות: לא נכתב דבר' };
        const cur = this.S.liveByKey.get(s.loc.key).tokens.slice(s.loc.from, s.loc.to + 1), tgt = splitWords(custom);
        const expect = tgt.map(canon).join(' ');
        if (!tgt.length) return { ok: false, why: 'הנוסח ריק' };
        if (expect !== cur.map(canon).join(' ')) {
          const edit = planTokenEdit(s.loc.key, s.loc.from, cur, tgt);
          if (!edit) return { ok: false, why: 'לא חושבה כתיבה' };
          edit.tag = id;
          const prevLog = this.log[id];
          this.log[id] = Object.assign({}, prevLog || {}, { state: 'other', span: expect, pending: true });
          const r = await this._write(edit);
          if (!r.ok) { if (prevLog) this.log[id] = prevLog; else delete this.log[id]; return r; }
          await this.refresh(it.pid);
          const s2 = this.state(id);
          if (!s2 || s2.current !== expect) { if (prevLog) this.log[id] = prevLog; else delete this.log[id]; this.S.compute(this.live, new Set([it.pid])); return { ok: false, failed: true, why: 'אחרי הכתיבה הטקסט אינו כמצופה. בטל ב-Ctrl+Z ובדוק', edit }; }
          delete this.log[id].pending;
        }
        this._log(id, { state: 'other', text: custom, span: expect });
        this.S.compute(this.live, new Set([it.pid]));
        return { ok: true };
      }
      if (!s || ['notfound', 'changed', 'na', 'blocked'].includes(s.state))
        return { ok: false, why: { notfound: 'המקום לא נמצא בוודאות', changed: 'המקום השתנה בקובץ', na: 'לא רלוונטי (פריט שהוא תלוי בו נדחה)', blocked: 'צריך לאשר קודם את הפריט שהוא תלוי בו' }[s && s.state] || 'אין מצב' };
      if (s.state === 'accepted' && custom === undefined) return { ok: true, already: true };
      let edit, expect;
      if (it.cluster && s.mask === undefined) return { ok: false, why: 'הנוסח במקום אינו המקור ואינו ההצעה (נערך בקובץ): לא נכתב דבר' };
      if (it.cluster) {
        const c = P.clusters.get(it.cluster), k = c.items.indexOf(id);
        const cur = this.S.liveByKey.get(s.loc.key).tokens.slice(s.loc.from, s.loc.to + 1);
        const mask = s.mask | (1 << k);
        const tgt = splitWords(outcomeText(c, P.items, mask, custom !== undefined ? { [id]: custom } : null));
        if (tgt.map(canon).join(' ') === cur.map(canon).join(' ')) return { ok: true, already: true };
        edit = planTokenEdit(s.loc.key, s.loc.from, cur, tgt);
        if (!edit) return { ok: false, why: 'לא חושבה כתיבה' };
        if (edit.kind !== 'tokenEdit' && edit.kind !== 'deleteTokens') edit.runs = pickRuns(it, edit.text, custom);
        expect = tgt.map(canon).join(' ');
        edit.tag = it.cluster;
        const prevLog = this.log[id];
        if (custom !== undefined) this.log[id] = Object.assign({}, prevLog || {}, { state: 'other', span: expect, pending: true });
        const restore = () => { if (custom !== undefined) { if (prevLog) this.log[id] = prevLog; else delete this.log[id]; } };
        const r = await this._write(edit);
        if (!r.ok) { restore(); return r; }
        await this.refresh(it.pid);
        const s2 = this.state(id);
        const got = s2 && s2.current;
        if (got !== expect) { restore(); return { ok: false, failed: true, why: 'אחרי הכתיבה הטקסט אינו כמצופה. בטל ב-Ctrl+Z ובדוק', edit }; }
        if (custom !== undefined) { delete this.log[id].pending; this._log(id, { state: 'other', text: custom, span: got }); }
        else if (this.log[id]) delete this.log[id];
        this.S.compute(this.live, new Set([it.pid]));
        return { ok: true, edit };
      }
      if (it.op === 'split_before') {
        edit = { kind: 'split', key: s.loc.key, token: s.loc.token, tag: id };
      } else {
        const text = custom !== undefined ? custom : it.new;
        edit = { kind: 'insertParagraphBefore', key: s.loc.key, text, style: it.para_style, runs: custom !== undefined ? null : it.runs, tag: id };
      }
      const r = await this._write(edit);
      if (!r.ok) return r;
      await this.refresh(it.pid);
      const s2 = this.state(id);
      if (custom !== undefined) {
        if (!s2 || !['changed', 'other', 'accepted'].includes(s2.state)) return { ok: false, failed: true, why: 'אחרי הכתיבה הפסקה לא נמצאה', edit };
        this._log(id, { state: 'other', text: custom, span: s2.current || custom }); this.S.compute(this.live);
      } else if (!s2 || s2.state !== 'accepted') return { ok: false, failed: true, why: 'אחרי הכתיבה המצב אינו "אושר". בטל ב-Ctrl+Z ובדוק', edit };
      return { ok: true, edit };
    }
    /* "אחר" רחב: יונתן כותב נוסח לטווח שכולל את מקום הפריט ומילים סביבו (key, from..to בפסקה החיה).
       נכתבות רק התיבות שהשתנו. אחר כך בודקים שהטווח בדיוק כמו שהוקלד. */
    async otherRange(id, key, from, to, text, why) {
      const it = this.P.items.get(id);
      if (!it) return { ok: false, why: 'אין פריט כזה' };
      await this.refresh(it.pid);
      const s = this.state(id);
      if (!s || !s.loc || s.loc.key !== key || s.loc.from < from || Math.max(s.loc.from, s.loc.to) > to) return { ok: false, why: 'המקום השתנה מאז שנפתח השדה: לא נכתב דבר' };
      const live = this.S.liveByKey.get(key);
      if (!live || from < 0 || to >= live.tokens.length) return { ok: false, why: 'הטווח לא נמצא' };
      const cur = live.tokens.slice(from, to + 1), tgt = splitWords(text);
      if (!tgt.length) return { ok: false, why: 'הנוסח ריק' };
      const expect = tgt.map(canon).join(' ');
      if (expect === cur.map(canon).join(' ')) return { ok: false, why: 'לא שונה דבר' };
      const edit = planTokenEdit(key, from, cur, tgt);
      if (!edit) return { ok: false, why: 'לא חושבה כתיבה' };
      edit.tag = 'wide:' + id;
      const r = await this._write(edit);
      if (!r.ok) return r;
      await this.refresh(it.pid);
      const after = (this.S.liveByKey.get(key) || { tokens: [] }).tokens.slice(from, from + tgt.length).map(canon).join(' ');
      if (after !== expect) return { ok: false, failed: true, why: 'אחרי הכתיבה הטקסט אינו כמצופה. בטל ב-Ctrl+Z ובדוק', edit };
      const s2 = this.state(id);
      this._log(id, { state: 'other', text, span: (s2 && s2.current) || null, wide: true, why: why || '' });
      this.S.compute(this.live, new Set([it.pid]));
      return { ok: true, edit };
    }
    async reject(id, why) { this._log(id, { state: 'rejected', why: why || '' }); return this.refresh(); }
    async markSeen(id) { this._log(id, { state: 'seen' }); return this.refresh(); }
    /* עריכה ישירה של יונתן במקום: לרשום כ"אחר" בלי לכתוב */
    async recordOther(id, why) {
      await this.refresh();
      const s = this.state(id);
      if (!s || s.state !== 'changed') return { ok: false, why: 'הפריט לא במצב "השתנה בקובץ"' };
      this._log(id, { state: 'other', text: s.current, span: s.current, why: why || '' });
      await this.refresh();
      return { ok: true };
    }
    async other(id, text, why) {
      const r = await this.approve(id, text);
      if (r.ok && why) this.log[id].why = why;
      return r;
    }
    _log(id, rec) { rec.time = new Date().toISOString(); this.log[id] = Object.assign({}, this.log[id] || {}, rec); if (this.opts.onLog) this.opts.onLog(id, this.log[id]); }
  }

  function pickRuns(it, text, custom) {
    const runs = (it.runs || []).filter((r) => r.text);
    if (!runs.length) return null;
    const styles = new Set(runs.map((r) => r.char_style || null));
    if (custom !== undefined) return styles.size === 1 ? [{ text, char_style: runs[0].char_style || null }] : null;
    const joined = runs.map((r) => r.text).join('');
    if (joined.trim() === text.trim()) {
      // חותכים רווחים בקצוות
      const out = runs.map((r) => ({ text: r.text, char_style: r.char_style || null }));
      while (out.length && !out[0].text.trim()) out.shift();
      if (out.length) out[0].text = out[0].text.replace(/^\s+/, '');
      while (out.length && !out[out.length - 1].text.trim()) out.pop();
      if (out.length) out[out.length - 1].text = out[out.length - 1].text.replace(/\s+$/, '');
      return out;
    }
    return styles.size === 1 ? [{ text, char_style: runs[0].char_style || null }] : null;
  }

  return { loadPackage, outcomeText, expectedParagraphs, matchParagraphs, locate, planTokenEdit, State, Session, norm, canon, splitWords, applyCharEdits };
}));
