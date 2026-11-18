/* בדיקות ה-port האמיתי (wordPortWord) מול דמה של Word (tests/wordMock.js). הרצה: node tests/port.test.js [--fuzz N]
   המסמך נבנה מ-orig של חבילה אמיתית, עם "רעש" של שינויים במעקב קיימים (מחיקות והוספות של אחרים) בתוך הפסקאות,
   וסימנים שכנים לפני ואחרי. בודקים: אישור כל הפריטים בסדר אקראי נותן את ההצעה; שום דבר מחוץ לסימן לא השתנה;
   הוספות של יונתן עצמו לא נמחקות לעולם; סגנונות כותרת ומ"מ. */
'use strict';
const C = require('../core.js'), { WordPort } = require('../wordPortWord.js'), { Doc, install, ME } = require('./wordMock.js');
const fs = require('fs'), path = require('path'), assert = require('assert');
const DIR = process.env.PKG_DIR || path.join(__dirname, '../../חבילות-לדוגמה');
const argv = process.argv.slice(2);
const FUZZ = +(argv[argv.indexOf('--fuzz') + 1] || 0) || 6;
let seed = +(process.env.SEED || 777);
const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
const shuffle = (a) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const W = (s) => C.splitWords(s).join(' ');
const pkgs = fs.readdirSync(DIR).filter((f) => f.startsWith('סקירה') && f.endsWith('.json') && !/קד/.test(f))
  .map((f) => ({ name: f.replace('סקירה - סימן ', '').replace('.json', ''), pkg: JSON.parse(fs.readFileSync(path.join(DIR, f), 'utf8')) }));

function build(pkg, noise, mine) {
  const d = new Doc();
  d.addPara('סימן קודם', '1_סימן'); d.addPara('טקסט של הסימן הקודם, לא לגעת.', '4_טקסט-רץ');
  for (const p of pkg.paragraphs) {
    const P = d.addPara(p.orig.replace(/\s+$/, ''), p.style);
    if (!noise) continue;
    // רעש: מילים מחוקות במעקב (של אחר) בין מילים, ותווים מחוקים בתוך מילים
    for (let k = P.chars.length - 1; k > 0; k--) {
      if (P.chars[k].c === ' ' && rnd() < 0.04) P.chars.splice(k, 0, ...[...' נמחק'].map((c) => ({ c, del: true, delBy: 'Chaim Rechnitzer', add: null, cstyle: null })));
      else if (rnd() < 0.01) P.chars.splice(k, 0, { c: ',', del: true, delBy: 'user', add: null, cstyle: null });
      else if (rnd() < 0.02) P.chars[k].add = 'Chaim Rechnitzer';
      else if (mine && rnd() < 0.01) { P.chars[k].add = ME; P.chars[k].mine0 = true; }
    }
  }
  d.addPara('סימן אחרי', '1_סימן'); d.addPara('טקסט של הסימן הבא.', '4_טקסט-רץ');
  return d;
}
const siman = (d) => { const i0 = d.paras.findIndex((p) => p.style === '1_סימן' && d.clean(p) !== 'סימן קודם'); let i1 = i0 + 1; while (d.clean(d.paras[i1]) !== 'סימן אחרי') i1++; return d.paras.slice(i0, i1); };
const outside = (d) => d.paras.filter((p) => !siman(d).includes(p)).map((p) => d.raw(p) + '|' + p.style).join('\n');

const results = [];
async function test(name, fn) {
  const t0 = Date.now();
  try { const info = await fn(); results.push(['✓', name, info || '']); } catch (e) { results.push(['✗', name, (e.stack || e.message).slice(0, 600)]); }
  const r = results[results.length - 1]; console.log(r[0] + ' ' + r[1] + (r[2] ? ' — ' + r[2] : '') + `  (${((Date.now() - t0) / 1000).toFixed(1)}s)`);
}
async function run(pkg, d, order) {
  install(d);
  const port = new WordPort(); const S = new C.Session(port, pkg, { log: {} });
  // ה-port מחפש "סימן <שם>": נותנים לשורת הסימן בדמה את השם
  await S.refresh();
  const ids = order || shuffle(pkg.items.filter((i) => i.op !== 'none').map((i) => i.id));
  let rest = ids; const fails = [];
  for (let k = 0; k < 5 && rest.length; k++) {
    const nx = [];
    for (const id of rest) { const r = await S.approve(id); if (r.ok) continue; if (r.failed || /לא נכתב דבר/.test(r.why || '')) fails.push(id + ': ' + r.why); else nx.push(id); }
    rest = nx;
  }
  return { S, port, rest, fails };
}

(async () => {
  for (const { name, pkg } of pkgs) {
    for (const noise of [false, true]) {
      await test(`${name}: אישור הכול ${noise ? 'עם שינויים קיימים במעקב' : 'במסמך נקי'} = ההצעה`, async () => {
        let info = '';
        for (let f = 0; f < (noise ? FUZZ : 1); f++) {
          const d = build(pkg, noise); const before = outside(d);
          const { rest, fails, port, S } = await run(pkg, d);
          if (rest.length) console.log(rest.map((id) => id + ' ' + JSON.stringify(S.state(id)) + ' ' + JSON.stringify(S.P.items.get(id).op)).join('\n'));
          assert.strictEqual(fails.length, 0, fails.slice(0, 4).join(' | ') + ' || ' + port.diag.slice(-3).join(' | '));
          // "לא נמצא" בסדר אקראי מסוים מותר (בטוח: לא נכתב דבר), עד 1% מהפריטים; אז לא משווים את הטקסט כולו
          const nf = rest.filter((id) => S.state(id).state === 'notfound');
          assert.ok(rest.length === nf.length && nf.length <= Math.ceil(pkg.items.length / 100), 'נשארו ' + rest.length);
          if (rest.length) { info = 'דילוג על השוואה: ' + rest.length + ' "לא נמצא"'; continue; }
          const got = siman(d).map((p) => W(d.clean(p))), want = pkg.paragraphs.flatMap((p) => p.proposed.map((x) => W(x.text)));
          const bad = got.findIndex((g, i) => g !== want[i]);
          assert.ok(got.length === want.length && bad < 0, `פסקה ${bad}: \n${got[bad]}\n${want[bad]}`);
          assert.strictEqual(outside(d), before, 'השתנה משהו מחוץ לסימן');
          // שום מחיקה של תו שנמחק קודם לא בוטלה
          for (const p of siman(d)) for (const x of p.chars) if (x.delBy && x.delBy !== ME) assert.ok(x.del);
          // כותרות ומ"מ
          const heads = siman(d).filter((p) => p.style === '7_כותרת-צד').length, wantH = pkg.paragraphs.flatMap((p) => p.proposed).filter((x) => x.style === '7_כותרת-צד').length;
          assert.strictEqual(heads, wantH, 'מספר כותרות צד');
          const mm = siman(d).flatMap((p) => p.chars).filter((x) => x.cstyle === '8_מ"מ' && x.add === ME).length;
          info = `כותרות ${heads}, תווי מ"מ שנכתבו ${mm}, ${port.diag.length} רישומים`;
        }
        return info;
      });
    }
    await test(`${name}: מיקומי מחיקות שגויים מ-Word: יישור חלופי, ועדיין = ההצעה`, async () => {
      const d = build(pkg, true); d.badTc = true; const before = outside(d);
      const { rest, fails, port } = await run(pkg, d);
      // מותר שפריט ייעצר בלי כתיבה ("לא אותר בוודאות"); אסור שייכתב משהו לא נכון
      const wrong = fails.filter((x) => !/לא אותר בוודאות/.test(x));
      assert.strictEqual(wrong.length + rest.length, 0, wrong.slice(0, 3).join(' | '));
      assert.ok(fails.length <= pkg.items.length * 0.05, 'יותר מדי פריטים נעצרו: ' + fails.length);
      if (!fails.length) { const got = siman(d).map((p) => W(d.clean(p))), want = pkg.paragraphs.flatMap((p) => p.proposed.map((x) => W(x.text))); assert.deepStrictEqual(got, want); }
      assert.strictEqual(outside(d), before);
      return port.diag.filter((x) => /יישור חלופי/.test(x)).length + ' יישורים חלופיים, ' + fails.length + ' נעצרו בלי כתיבה';
    });
    await test(`${name}: פסיק ליד מילה שיונתן הוסיף בעצמו נכתב`, async () => {
      const d = build(pkg, false); install(d); const port = new WordPort(); const S = new C.Session(port, pkg, { log: {} });
      await S.refresh();
      const it = pkg.items.find((i) => i.cluster && i.kind === 'punct' && i.op === 'suffix' && S.state(i.id).state === 'open');
      if (!it) return 'אין פריט מתאים';
      const s0 = S.state(it.id); const P = d.paras.find((x) => x.id === s0.loc.key);
      // יונתן מוסיף "(??)" אחרי האשכול, במעקב
      const ce = (() => { const toks = [...d.clean(P).matchAll(/\S+/g)]; const T = toks[s0.loc.from]; return T.index + T[0].length; })();
      const rawPos = (() => { let n = 0; for (let k = 0; k < P.chars.length; k++) if (!P.chars[k].del && ++n === ce) return k + 1; })();
      P.chars.splice(rawPos, 0, ...[...' (??)'].map((c) => ({ c, del: false, add: ME, cstyle: null })));
      await S.refresh(); const st1 = S.state(it.id).state;
      const r = await S.approve(it.id);
      assert.ok(r.ok, (r.why || '') + ' מצב לפני: ' + st1 + ' | ' + port.diag.slice(-2).join(' | '));
      assert.ok(d.clean(P).includes(String(it.new).trim()), 'הסימן לא נכתב');
      return 'מצב לפני ' + st1 + ', אחרי ' + S.state(it.id).state;
    });
    await test(`${name}: הוספות של יונתן עצמו לא נמחקות`, async () => {
      let refused = 0;
      for (let f = 0; f < FUZZ; f++) {
        const d = build(pkg, true, true);
        const { fails } = await run(pkg, d);
        assert.ok(!d.lostMine, 'נמחקה הוספה של יונתן');
        refused += fails.filter((x) => /הוספה שלך/.test(x)).length;
        const other = fails.filter((x) => !/הוספה שלך/.test(x));
        assert.strictEqual(other.length, 0, other.slice(0, 3).join(' | '));
      }
      return `סירובים בגלל הוספה שלו: ${refused}`;
    });
    await test(`${name}: מעקב כבוי = לא נכתב דבר`, async () => {
      const d = build(pkg, false); install(d); const port = new WordPort(); const S = new C.Session(port, pkg, { log: {} });
      await S.refresh(); d.mode = 'Off'; port.setTrackingOn = async () => false;
      const it = pkg.items.find((i) => ['suffix', 'replace', 'insert_after'].includes(i.op)); if (!it) return 'אין פריט מתאים'; const r = await S.approve(it.id);
      assert.ok(!r.ok && d.ops === 0, 'נכתב בלי מעקב');
    });
    await test(`${name}: סימון ב-Word מסמן את המילים הנכונות גם בפסקה עם מחיקות`, async () => {
      const d = build(pkg, true); const { S, port } = await run(pkg, d, []);
      const its = pkg.items.filter((i) => i.op === 'replace' && i.old.trim() && S.state(i.id).loc).slice(0, 20);
      let n = 0;
      for (const it of its) {
        const s = S.state(it.id); await port.mark(s.loc.key, s.loc.from, s.loc.to);
        const ids = port.markIds || []; assert.strictEqual(ids.length, 1, 'לא סומן: ' + port.diag.slice(-1));
        const r = d.annots[ids[0]]; const P = r.p0;
        const got = P.chars.slice(r.a, r.b).filter((x) => !x.del).map((x) => x.c).join('');
        const want = S.live.find((q) => q.key === s.loc.key).tokens.slice(s.loc.from, s.loc.to + 1).join(' ');
        assert.strictEqual(got.replace(/\s+/g, ' '), want, 'סומן מקום אחר'); n++;
      }
      await port.mark(null); assert.strictEqual(Object.keys(d.annots || {}).length, 0, 'הסימון הקודם לא נמחק');
      assert.strictEqual(d.ops, 0, 'סימון כתב משהו');
      return n + ' סימונים';
    });
    await test(`${name}: מראה מקום שנכתב בתוך תיבה מקבל סגנון מ"מ`, async () => {
      const d = build(pkg, false); const { S } = await run(pkg, d);
      const big = (r) => r.char_style === '8_מ"מ' && r.text.trim().length >= 4;
      const refs = pkg.items.filter((i) => i.kind === 'ref' && (i.runs || []).some(big));
      let ok = 0, bad = [];
      for (const it of refs) {
        const s = S.state(it.id); if (!s || s.state !== 'accepted' || !s.loc) continue;
        const P = d.paras.find((x) => x.id === s.loc.key); if (!P) continue;
        const vis = P.chars.filter((x) => !x.del), txt = vis.map((x) => x.c).join('');
        for (const r of it.runs.filter(big)) {
          const k = txt.indexOf(r.text.trim()); if (k < 0) continue;
          if (vis.slice(k, k + r.text.trim().length).every((x) => x.cstyle === '8_מ"מ')) ok++; else bad.push(it.id + ' ' + r.text.trim());
        }
      }
      assert.strictEqual(bad.length, 0, bad.slice(0, 5).join(' | '));
      return ok + ' מראי מקום';
    });
    await test(`${name}: בחירה ובחירה רחבה`, async () => {
      const d = build(pkg, true); const { S, port } = await run(pkg, d, []);
      const it = pkg.items.find((i) => i.op === 'replace' && i.old.trim()); if (!it) return 'אין פריט מתאים'; const s = S.state(it.id);
      await port.select(s.loc.key, Math.max(0, s.loc.from - 2), s.loc.to + 2);
      const t = d.sel.text || d.raw(d.sel.p0).slice(d.sel.a, d.sel.b);
      assert.ok(t.includes(C.splitWords(it.old)[0] || ''), 'הבחירה אינה במקום');
      assert.strictEqual(d.ops, 0, 'בחירה כתבה משהו');
    });
  }
  const bad = results.filter((r) => r[0] === '✗').length;
  console.log(`\n${results.length - bad}/${results.length} עברו`);
  process.exit(bad ? 1 : 0);
})();
