/* בדיקות הליבה (שלב 2). הרצה: node tests/core.test.js [--fuzz N] [--quick]
   על חבילות אמיתיות מ-ממשק-סקירה/חבילות-לדוגמה, מול מסמך דמה שנבנה מהמקור.
   הבדיקה החשובה: אף כתיבה לא נוחתת במקום שגוי (wordPortFake בודק את מקור כל תיבה שנכתבה). */
'use strict';
const C = require('../core.js'), { FakeDoc } = require('../wordPortFake.js');
const fs = require('fs'), path = require('path'), assert = require('assert');
const DIR = process.env.PKG_DIR || path.join(__dirname, '../../חבילות-לדוגמה');
const argv = process.argv.slice(2);
const FUZZ = +(argv[argv.indexOf('--fuzz') + 1] || 0) || (argv.includes('--quick') ? 40 : 300);
let seed = +(process.env.SEED || 12345);
const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
const pick = (a) => a[Math.floor(rnd() * a.length)];
const shuffle = (a) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const W = (s) => C.splitWords(s).join(' ');

const pkgs = fs.readdirSync(DIR).filter((f) => f.startsWith('סקירה') && f.endsWith('.json'))
  .map((f) => ({ name: f.replace('סקירה - סימן ', '').replace('.json', ''), pkg: JSON.parse(fs.readFileSync(path.join(DIR, f), 'utf8')) }));
const fresh = (pkg, log) => { const doc = FakeDoc.fromPackage(pkg); const S = new C.Session(doc, pkg, { log: log || {} }); doc.setPackage(S.P); return { doc, S }; };
const writable = (pkg) => pkg.items.filter((i) => i.op !== 'none').map((i) => i.id);
const proposedAll = (pkg) => pkg.paragraphs.flatMap((p) => p.proposed.map((x) => W(x.text)));
const origAll = (pkg) => pkg.paragraphs.map((p) => W(p.orig));

async function approveAll(S, ids, passes = 5) {
  let rest = ids, okc = 0;
  for (let k = 0; k < passes && rest.length; k++) {
    const nx = [];
    for (const id of rest) {
      const r = await S.approve(id); if (r.ok) okc++; else if (!r.failed) nx.push(id);
      else {
        if (process.env.DEBUG) { const st = S.state(id), it = S.P.items.get(id); const M = S.S.match;
          console.log('FAILDBG', JSON.stringify({ it: { op: it.op, pid: it.pid, from: it.from, new: it.new.slice(0, 40) }, st, edit: r.edit && { ...r.edit, runs: 0 } }));
          const live = S.live; const i = live.findIndex((x) => x.key === r.edit.key);
          live.slice(Math.max(0, i - 3), i + 2).forEach((x) => console.log('  ', x.key, x.style, x.tokens.slice(0, 8).join(' ')));
          console.log('  match', JSON.stringify(M.map.get(it.pid))); }
        throw new Error(id + ': ' + r.why);
      }
    }
    rest = nx;
  }
  return rest;
}

const results = [];
async function test(name, fn) {
  const t0 = Date.now();
  try { const info = await fn(); results.push(['✓', name, info || '', Date.now() - t0]); }
  catch (e) { results.push(['✗', name, e.message.slice(0, 400), Date.now() - t0]); }
  const r = results[results.length - 1]; console.log(r[0] + ' ' + r[1] + (r[2] ? ' — ' + r[2] : '') + `  (${(r[3] / 1000).toFixed(1)}s)`);
}

(async () => {
  const ONLY_FUZZ = argv.includes('--only-fuzz');
  for (const { name, pkg } of (ONLY_FUZZ ? [] : pkgs)) {
    const big = pkg.items.length > 300;

    await test(`${name}: בטעינה כל הפריטים נמצאים ופתוחים`, async () => {
      const { S } = fresh(pkg); const st = await S.refresh();
      const bad = [...st].filter(([, v]) => !['open', 'info', 'blocked'].includes(v.state));
      assert.strictEqual(bad.length, 0, bad.slice(0, 5).map(([k, v]) => k + ':' + v.state + ' ' + (v.why || '')).join('; '));
      const bl = [...st].filter(([k, v]) => v.state === 'blocked' && !(S.P.items.get(k).depends_on || []).length);
      assert.strictEqual(bl.length, 0, 'חסום בלי תלות');
      return `${st.size} פריטים`;
    });

    for (let r = 0; r < (big ? 1 : 3); r++) await test(`${name}: אישור כל הפריטים בסדר אקראי (${r + 1}) = ההצעה`, async () => {
      const { doc, S } = fresh(pkg);
      await S.refresh();
      const left = await approveAll(S, shuffle(writable(pkg)));
      assert.strictEqual(doc.violations.length, 0, 'כתיבה במקום שגוי: ' + JSON.stringify(doc.violations[0]));
      assert.deepStrictEqual(left, [], 'לא אושרו: ' + left.slice(0, 5));
      assert.deepStrictEqual(doc.dump().map((x) => x.text), proposedAll(pkg));
      const st = await S.refresh();
      const notacc = [...st].filter(([, v]) => !['accepted', 'info'].includes(v.state));
      assert.strictEqual(notacc.length, 0, notacc.slice(0, 3).map(([k, v]) => k + ':' + v.state).join(' '));
      return `${doc.revs.length} כתיבות`;
    });

    await test(`${name}: דחיית הכול לא משנה דבר`, async () => {
      const { doc, S } = fresh(pkg); await S.refresh();
      for (const id of writable(pkg)) await S.reject(id);
      assert.strictEqual(doc.revs.length, 0);
      assert.deepStrictEqual(doc.dump().map((x) => x.text), origAll(pkg));
      const st = S.S.st; assert.ok(writable(pkg).every((id) => ['rejected', 'na'].includes(st.get(id).state)));
    });

    await test(`${name}: כל פריט לבדו נותן רק את השינוי שלו`, async () => {
      const ids = big ? shuffle(writable(pkg)).slice(0, 60) : writable(pkg);
      let n = 0;
      for (const id of ids) {
        const it = pkg.items.find((x) => x.id === id);
        if ((it.depends_on || []).length) continue;
        const { doc, S } = fresh(pkg); await S.refresh();
        const r = await S.approve(id);
        assert.ok(r.ok, id + ': ' + r.why);
        assert.strictEqual(doc.violations.length, 0, id + ' במקום שגוי');
        const exp = C.expectedParagraphs(S.P, it.pid, new Set([id])).map((x) => x.tokens.join(' '));
        const all = []; for (const p of pkg.paragraphs) { if (p.pid === it.pid) all.push(...exp); else all.push(W(p.orig)); }
        assert.deepStrictEqual(doc.dump().map((x) => x.text), all, id);
        const st = S.S.st;
        for (const j of S.P.paras.get(it.pid).items) if (j !== id && S.P.items.get(j).op !== 'none')
          assert.ok(['open', 'blocked'].includes(st.get(j).state), `${id} שינה את מצב ${j}: ${st.get(j).state}`);
        n++;
      }
      return `${n} פריטים`;
    });

    await test(`${name}: באשכולות של כמה פריטים, כל צירוף של אישור ודחייה`, async () => {
      let n = 0;
      const cls = pkg.clusters.filter((c) => c.items.length > 1);
      for (const c0 of (big ? shuffle(cls).slice(0, 25) : cls)) {
        const k = c0.items.length;
        for (let m = 1; m < (1 << k); m++) {
          const { doc, S } = fresh(pkg); await S.refresh();
          const sub = c0.items.filter((_, i) => m >> i & 1);
          for (const id of shuffle(sub)) { const r = await S.approve(id); assert.ok(r.ok, id + ' ' + r.why); }
          for (const id of c0.items.filter((x) => !sub.includes(x))) await S.reject(id);
          assert.strictEqual(doc.violations.length, 0);
          const c = S.P.clusters.get(c0.cid);
          const st = S.S.st.get(c.items[0]);
          assert.strictEqual(st.current, c.outKeys[m], c0.cid + ' מסכה ' + m);
          c.items.forEach((id, i) => assert.strictEqual(S.S.st.get(id).state, (m >> i & 1) ? 'accepted' : 'rejected', id));
          n++;
        }
      }
      return `${n} צירופים ב-${cls.length} אשכולות`;
    });
  }

  // ───── "אחר", עריכה ישירה, מעקב כבוי (על קה)
  if (!ONLY_FUZZ) await (async () => {
  const small = pkgs.find((x) => x.pkg.siman === 'קה') || pkgs.find((x) => x.pkg.items.length < 200) || pkgs[0];
  await test(`${small.name}: "אחר" נכתב במקום ההצעה, ונשמר אחרי טעינה מחדש`, async () => {
    const { doc, S } = fresh(small.pkg); await S.refresh();
    const word = small.pkg.items.find((i) => i.cluster && ['word', 'spelling', 'abbr', 'ref'].includes(i.kind) && i.new.trim());
    const punct = small.pkg.items.find((i) => i.cluster && i.kind === 'punct' && i.new === ',');
    const head = small.pkg.items.find((i) => i.op === 'heading_before' && !(i.depends_on || []).length);
    const r1 = await S.other(word.id, word.new.trim() + 'א', 'בדיקה');
    assert.ok(r1.ok, r1.why); assert.strictEqual(S.state(word.id).state, 'other');
    const r2 = await S.other(punct.id, ';'); assert.ok(r2.ok, r2.why); assert.strictEqual(S.state(punct.id).state, 'other');
    if (head) { const r3 = await S.other(head.id, 'כותרת בנוסח של יונתן'); assert.ok(r3.ok, r3.why); assert.strictEqual(S.state(head.id).state, 'other'); }
    assert.strictEqual(doc.violations.length, 0);
    // טעינה מחדש: אותו מסמך, יומן שמור
    const S2 = new C.Session(doc, small.pkg, { log: JSON.parse(JSON.stringify(S.log)) });
    await S2.refresh();
    assert.strictEqual(S2.state(word.id).state, 'other'); assert.strictEqual(S2.state(punct.id).state, 'other');
    if (head) assert.strictEqual(S2.state(head.id).state, 'other');
    // בלי יומן: "השתנה בקובץ"
    const S3 = new C.Session(doc, small.pkg); await S3.refresh();
    assert.strictEqual(S3.state(word.id).state, 'changed');
    return 'מילה, פיסוק' + (head ? ', כותרת' : '');
  });

  await test(`${small.name}: עריכה ישירה במסמך = "השתנה בקובץ", ואז נרשם כ"אחר" בלי כתיבה`, async () => {
    const { doc, S } = fresh(small.pkg); await S.refresh();
    const it = small.pkg.items.find((i) => i.cluster && i.kind === 'word' && i.op === 'replace') || small.pkg.items.find((i) => i.cluster && i.kind !== 'punct');
    const s = S.state(it.id);
    doc.userEditToken(s.loc.key, s.loc.from, 'נוסחיונתן');
    await S.refresh(); assert.strictEqual(S.state(it.id).state, 'changed');
    const nrev = doc.revs.length;
    const r = await S.recordOther(it.id, 'עריכה ישירה'); assert.ok(r.ok, r.why);
    assert.strictEqual(S.state(it.id).state, 'other'); assert.strictEqual(doc.revs.length, nrev);
    const r2 = await S.approve(it.id); assert.ok(!r2.ok, 'אישור אחרי "השתנה" לא אמור לכתוב');
  });

  await test(`${small.name}: הקלדת ההצעה ביד = "אושר"`, async () => {
    const { doc, S } = fresh(small.pkg); await S.refresh();
    const it = small.pkg.items.find((i) => i.cluster && i.op === 'suffix');
    const s = S.state(it.id); const cur = doc._clean(doc._para(s.loc.key))[s.loc.from].t;
    doc.userEditToken(s.loc.key, s.loc.from, cur + it.new);
    await S.refresh(); assert.strictEqual(S.state(it.id).state, 'accepted');
  });

  await test(`${small.name}: מעקב כבוי שלא נדלק = לא נכתב דבר`, async () => {
    const { doc, S } = fresh(small.pkg); await S.refresh();
    doc.tracking = false; doc.lockTracking = true;
    const r = await S.approve(writable(small.pkg)[0]);
    assert.ok(!r.ok && /מעקב/.test(r.why)); assert.strictEqual(doc.revs.length, 0);
    doc.lockTracking = false; const r2 = await S.approve(writable(small.pkg)[0]); assert.ok(r2.ok); assert.ok(doc.tracking);
  });

  await test(`${small.name}: כותרת לפסקה שהפיצול שלה נדחה = לא רלוונטי`, async () => {
    const { S } = fresh(small.pkg); await S.refresh();
    const h = small.pkg.items.find((i) => (i.depends_on || []).length);
    if (!h) return 'אין';
    await S.reject(h.depends_on[0]);
    assert.strictEqual(S.state(h.id).state, 'na');
    const r = await S.approve(h.id); assert.ok(!r.ok);
  });

  })();
  // ───── הגרלות: אישור חלקי, עריכות של יונתן, המשך אישור
  await test(`הגרלות עריכה (${FUZZ}): אף כתיבה לא במקום שגוי`, async () => {
    const stat = { runs: 0, writes: 0, edits: 0, untouchedLost: 0, untouched: 0, notfound: 0 };
    const pool = pkgs.filter((x) => x.pkg.items.length <= 300);
    const heavy = pkgs.filter((x) => x.pkg.items.length > 300);
    for (let run = 0; run < FUZZ; run++) {
      const { pkg } = (heavy.length && run % 15 === 14) ? pick(heavy) : pick(pool);
      const { doc, S } = fresh(pkg); await S.refresh();
      const ids = shuffle(writable(pkg));
      const cut = Math.floor(ids.length * rnd() * 0.6);
      await approveAll(S, ids.slice(0, cut), 2);
      // עריכות של יונתן, מרוכזות בפסקה אחת או שתיים
      const nEd = 3 + Math.floor(rnd() * 12);
      const targets = shuffle(doc.paras.filter((p) => doc._clean(p).length > 6)).slice(0, 2);
      for (let e = 0; e < nEd; e++) {
        const p = pick(targets); const n = doc._clean(p).length; if (n < 4) continue;
        const i = 1 + Math.floor(rnd() * (n - 2)); const kind = Math.floor(rnd() * 7);
        const t = doc._clean(p)[i].t;
        if (kind === 0) doc.userDeleteToken(p.key, i);
        else if (kind === 1) doc.userInsertToken(p.key, i, pick(['ועוד', 'דהנה', 'וכו\'', 'בזה', 'שם']));
        else if (kind === 2) doc.userEditToken(p.key, i, /[.,:;]$/.test(t) ? t.slice(0, -1) : t + ',');
        else if (kind === 3) doc.userEditToken(p.key, i, t.length > 2 ? t.slice(0, -2) + 'ו' + t.slice(-2) : t + 'ה');
        else if (kind === 4 && i > 1) doc.userSplit(p.key, i);
        else if (kind === 5) doc.userAddParagraph(p.key, 'פסקה חדשה של יונתן עם כמה מילים', p.style);
        else doc.userEditToken(p.key, i, t);   // נגיעה בלי שינוי
        stat.edits++;
      }
      await S.refresh();
      const before = doc.revs.length;
      const left = await approveAll(S, shuffle(ids.slice(cut)), 3);
      stat.writes += doc.revs.length - before;
      if (doc.violations.length) {
        const v = doc.violations[0], c = S.P.clusters.get(v.edit.tag);
        if (c) console.log('DEBUG', JSON.stringify({ orig: c.orig, from: c.from, cb: c.ctx_before, ca: c.ctx_after, outs: c.outKeys }), '\nLIVE', doc.text(v.edit.key).split(' ').slice(Math.max(0, v.edit.token - 8), v.edit.token + 6).join(' '),
          '\nORIG', S.P.paras.get(c.pid).orig_tokens.slice(c.from - 8, c.to + 6).join(' '));
        throw new Error(`הרצה ${run} (${pkg.siman}): כתיבה במקום שגוי: ` + JSON.stringify(v).slice(0, 300));
      }
      // חיוּת: פריט שהתיבות שלו וההקשר שלו לא נגעו בהם צריך להימצא
      await S.refresh();
      for (const id of left) {
        const st = S.state(id); if (st.state !== 'notfound') continue;
        stat.notfound++;
        const it = S.P.items.get(id);
        const mine = doc.paras.filter((p) => p.toks.some((x) => x.origin && x.origin.pid === it.pid));
        const touched = mine.some((p) => p.toks.some((x) => x.touched)) || mine.length !== S.S.match.map.get(it.pid).keys.length;
        stat.untouched += !touched; if (!touched) stat.untouchedLost++;
      }
      stat.runs++;
    }
    const lostTxt = `${stat.notfound} פריטים "לא נמצא" אחרי עריכות, מהם ${stat.untouchedLost} בלי עריכה בקרבתם`;
    assert.ok(stat.untouchedLost <= Math.max(2, stat.runs * 0.02), lostTxt);
    return `${stat.runs} הרצות, ${stat.edits} עריכות של יונתן, ${stat.writes} כתיבות, 0 במקום שגוי. ${lostTxt}`;
  });

  const bad = results.filter((r) => r[0] === '✗').length;
  console.log(`\n${results.length - bad} מתוך ${results.length} בדיקות עברו` + (bad ? `, ${bad} נכשלו` : ''));
  process.exit(bad ? 1 : 0);
})();
