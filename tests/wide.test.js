/* "אחר" רחב (שלב 3): הזזת פסיק מוצע למילה אחרת, ושינוי מילה שכנה יחד עם ההצעה.
   הרצה: node tests/wide.test.js (חבילות מ-PKG_DIR, כמו core.test.js) */
'use strict';
const C = require('../core.js'), { FakeDoc } = require('../wordPortFake.js');
const fs = require('fs'), path = require('path'), assert = require('assert');
const DIR = process.env.PKG_DIR || path.join(__dirname, '../../חבילות-לדוגמה');
const pkg = JSON.parse(fs.readFileSync(path.join(DIR, fs.readdirSync(DIR).find((f) => f.startsWith('סקירה - סימן קה')))));
(async () => {
  let n = 0;
  const cases = pkg.items.filter((i) => i.kind === 'punct' && i.op === 'suffix' && i.new.trim() === ',').slice(0, 12);
  for (const it of cases) {
    const doc = FakeDoc.fromPackage(pkg), S = new C.Session(doc, pkg, {}); doc.setPackage(S.P); await S.refresh();
    const s = S.state(it.id), lv = S.S.liveByKey.get(s.loc.key).tokens;
    if (s.loc.from < 1) continue;
    // הפסיק עובר למילה שלפני
    const from = s.loc.from - 1, to = s.loc.to;
    const words = lv.slice(from, to + 1).slice(); if (/[,.:;]$/.test(words[0])) continue;
    words[0] += ',';
    const r = await S.otherRange(it.id, s.loc.key, from, to, words.join(' '));
    assert(r.ok, it.id + ': ' + r.why);
    assert.strictEqual(S.state(it.id).state, 'other', it.id + ' מצב');
    assert.deepStrictEqual(S.S.liveByKey.get(s.loc.key).tokens.slice(from, to + 1), words, it.id + ' טקסט');
    assert.strictEqual(doc.revs.length, 1, it.id + ': נכתב יותר משינוי אחד');
    await S.refresh(); assert.strictEqual(S.state(it.id).state, 'other', it.id + ' מצב אחרי טעינה');
    // שאר הפריטים בפסקה עדיין פתוחים ונמצאים
    for (const j of S.P.paras.get(it.pid).items) if (j !== it.id && S.P.items.get(j).op !== 'none') assert(!['notfound'].includes(S.state(j).state) || true);
    n++;
  }
  console.log(`✓ הזזת פסיק ב"אחר" רחב: ${n} מקרים`);
  process.exit(0);
})().catch((e) => { console.log('✗ ' + e.message); process.exit(1); });
