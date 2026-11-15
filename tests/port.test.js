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
const siman = (d) => { const i0 = d.paras.findIndex((p) => p.style === '1_סימן' && d.clean(p) !== 'סימן קודם'); let i1 = i0 + 1; while (d.paras[i1].style !== '1_סימן') i1++; return d.paras.slice(i0, i1); };
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
          const { rest, fails, port } = await run(pkg, d);
          assert.strictEqual(fails.length, 0, fails.slice(0, 4).join(' | ') + ' || ' + port.diag.slice(-3).join(' | '));
          assert.strictEqual(rest.length, 0, 'נשארו ' + rest.length);
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
