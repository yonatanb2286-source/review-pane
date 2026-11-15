/* בדיקות החלונית (שלב 3), Playwright מול מצב ההדגמה (wordPortFake).
   הרצה: node tests/pane.test.js [--shots DIR]
   החבילות נקראות מ-PKG_DIR (ברירת מחדל ../../חבילות-לדוגמה). הן מכילות טקסט מהספר: לא עולות ל-GitHub, וגם הצילומים לא.
   הבדיקה המרכזית: אחרי אישור, דחייה או "אחר" אף מילה בטקסט היציב לא זזה אפילו בפיקסל, כל עוד נשארים באותו קטע. */
'use strict';
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const ROOT = path.resolve(__dirname, '..');
const DIR = process.env.PKG_DIR || path.join(ROOT, '../חבילות-לדוגמה');
const argv = process.argv.slice(2);
const SHOTS = argv.includes('--shots') ? argv[argv.indexOf('--shots') + 1] : null;
const SIZES = [[329, 600], [350, 700]];
const pkgFile = (s) => fs.readdirSync(DIR).find((f) => f.startsWith('סקירה - סימן ' + s + ' '));
const PKGS = fs.readdirSync(DIR).filter((f) => f.startsWith('סקירה') && f.endsWith('.json'));
let pass = 0, fail = 0;
const ok = (c, msg) => { if (c) pass++; else { fail++; console.log('  ✗ ' + msg); } };

let browser;
async function open(file, [w, h], dsf) {
  const ctx = await browser.newContext({ viewport: { width: w, height: h }, deviceScaleFactor: dsf || 1, acceptDownloads: true });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => { fail++; console.log('  ✗ שגיאת דף: ' + e.message); });
  await p.addInitScript('window.DEMO_PKG=' + fs.readFileSync(path.join(DIR, file), 'utf8') + ';');
  await p.goto('file://' + path.join(ROOT, 'review.html'));
  await p.waitForFunction(() => window.__pane && __pane.cur);
  await p.evaluate(() => document.fonts.ready);
  return p;
}
const snap = (p) => p.evaluate(() => {
  const r = (e) => { const b = e.getBoundingClientRect(); return [b.left, b.top, b.width, b.height].map((x) => Math.round(x * 100) / 100).join(','); };
  return { where: __pane.pid + ':' + __pane.chunk, cur: __pane.cur,
    words: [...document.querySelectorAll('#text .w, #text .it')].map((e) => (e.dataset.id || e.textContent) + '@' + r(e)),
    regions: ['status', 'text', 'det', 'btns', 'row1', 'bPrev', 'bNext', 'bNote'].map((id) => id + '@' + r(document.getElementById(id))),
    btns: [...document.getElementById('row1').children].map((e) => e.className + '@' + r(e)) };
});
const noScroll = (p) => p.evaluate(() => {
  const d = document.documentElement, t = document.getElementById('text');
  return d.scrollWidth <= innerWidth && d.scrollHeight <= innerHeight && t.scrollHeight <= t.clientHeight + 1 && t.scrollWidth <= t.clientWidth + 1;
});
// פריט פתוח מהסוג, שהפריט הפתוח הבא אחריו באותו קטע (כדי שהטקסט יישאר מוצג)
const findCase = (p, kind, skip) => p.evaluate(([kind, skip]) => {
  const A = __pane, NEED = ['open', 'blocked', 'changed', 'notfound', 'info'];
  const st = (id) => (A.S.state(id) || {}).state;
  for (const id of A.order) {
    const it = A.P.items.get(id);
    if (it.kind !== kind || st(id) !== 'open' || skip.includes(id)) continue;
    const i = A.order.indexOf(id);
    const nx = A.order.slice(i + 1).find((x) => NEED.includes(st(x)));
    if (nx && A.P.items.get(nx).pid === it.pid) return [id, nx];
  }
  return null;
}, [kind, skip || []]);

async function stability(file) {
  const p = await open(file, SIZES[0]);
  let n = 0;
  for (const kind of ['punct', 'word', 'spelling', 'abbr', 'ref', 'heading', 'split', 'translation']) {
    for (const act of ['approve', 'reject', 'other']) {
      if (act === 'other' && kind === 'split') continue;
      const c = await findCase(p, kind); if (!c) continue;
      const [id, nx] = c;
      await p.evaluate((id) => __pane.go(id), id);
      const nxSame = await p.evaluate(([id, nx]) => { const A = __pane, [a, b] = A.chunks[A.chunk], u = A.unitOf.get(nx); return u >= a && u <= b; }, [id, nx]);
      const before = await snap(p);
      if (act === 'approve') await p.click('#bOk');
      else if (act === 'reject') await p.click('#bNo');
      else {
        await p.click('#bOt');
        const mid = await snap(p);
        ok(JSON.stringify(mid.words) === JSON.stringify(before.words) && JSON.stringify(mid.regions) === JSON.stringify(before.regions), `${file} ${id}: פתיחת "אחר" הזיזה משהו`);
        const v = await p.evaluate((id) => { const it = __pane.P.items.get(id); const s = it.new.trim(); return /^[.,:;]$/.test(s) ? (s === ';' ? ':' : ';') : s + 'ו'; }, id);
        await p.fill('#othT', v); await p.click('#othOk');
      }
      await p.waitForTimeout(60); await p.waitForFunction(() => !__pane.busy);
      const em = await p.evaluate(() => __pane.msg); if (em) { ok(false, `${file} ${id} ${act}: ${em}`); continue; }
      const after = await snap(p);
      const want = { approve: ['accepted'], reject: ['rejected'], other: ['other'] }[act];
      const s = await p.evaluate((id) => __pane.S.state(id).state, id);
      ok(want.includes(s), `${file} ${id} ${act}: המצב ${s}`);
      const nx2 = await p.evaluate((id) => { const A = __pane, N = ['open', 'blocked', 'changed', 'notfound', 'info'], i = A.order.indexOf(id);
        return A.order.slice(i + 1).concat(A.order.slice(0, i)).find((x) => N.includes(A.S.state(x).state)); }, id);
      ok(after.cur === nx2, `${file} ${id}: ההדגשה לא עברה לפריט הפתוח הבא (${after.cur} במקום ${nx2})`);
      ok(JSON.stringify(after.regions) === JSON.stringify(before.regions), `${file} ${id} ${act}: אזור זז או שינה גודל`);
      if (after.btns.length === before.btns.length) ok(JSON.stringify(after.btns) === JSON.stringify(before.btns), `${file} ${id} ${act}: כפתור זז`);
      if (nxSame && nx2 === nx) {
        ok(after.where === before.where, `${file} ${id} ${act}: הקטע התחלף אף שהפריט הבא בתוכו`);
        const moved = before.words.filter((w, i) => w !== after.words[i]);
        ok(!moved.length && before.words.length === after.words.length, `${file} ${id} ${kind} ${act}: ${moved.length} מילים זזו, למשל ${moved.slice(0, 2).join(' | ')}`);
        n++;
      }
      ok(await noScroll(p), `${file} ${id}: יש גלילה במסך הראשי`);
    }
  }
  await p.context().close();
  return n;
}

// הליכה על פסקה ארוכה באישורים ברצף: החלוקה לקטעים לא משתנה, והקטע מתחלף רק כשצריך
async function walk(file, steps) {
  const p = await open(file, SIZES[0]);
  const res = await p.evaluate(async (steps) => {
    const A = __pane, errs = [];
    let chunksOf = {};
    for (let k = 0; k < steps; k++) {
      const pid = A.pid, ch = A.chunk, rng = A.chunks[ch], cj = JSON.stringify(A.chunks);
      if (chunksOf[pid] && chunksOf[pid] !== cj) errs.push('החלוקה השתנתה בפסקה ' + pid);
      chunksOf[pid] = cj;
      const t = document.getElementById('text');
      if (t.scrollHeight > t.clientHeight + 1) errs.push('גלילה בקטע ' + pid + ':' + ch);
      const it = A.P.items.get(A.cur), st = A.S.state(A.cur).state;
      document.dispatchEvent(new KeyboardEvent('keydown', { key: st === 'open' ? 'Enter' : (it.op === 'none' ? 'Enter' : 'ArrowLeft'), bubbles: true }));
      for (let w = 0; w < 50 && (A.busy || A.cur === it.id); w++) await new Promise((r) => setTimeout(r, 10));
      if (A.pid === pid) {
        const u = A.unitOf.get(A.cur), inOld = u >= rng[0] && u <= rng[1];
        if (inOld && A.chunk !== ch) errs.push('הקטע התחלף בלי צורך ב-' + A.cur);
        if (!inOld && A.chunk === ch) errs.push('הפריט ' + A.cur + ' מחוץ לקטע המוצג');
      }
    }
    return { errs, left: document.getElementById('sLeft').textContent };
  }, steps);
  res.errs.forEach((e) => ok(false, file + ': ' + e));
  ok(!res.errs.length, file + ': הליכה');
  await p.context().close();
  return res.left;
}

// כל פריט בחבילה (חוץ מטכני) מסומן בטקסט, ופסיק מוצע רחב מספיק
async function marks(file) {
  const p = await open(file, SIZES[0]);
  const r = await p.evaluate(() => {
    const A = __pane, miss = [];
    const pids = [...new Set(A.order.map((id) => A.P.items.get(id).pid))];
    let narrow = 0, checked = 0;
    for (const pid of pids) {
      A.go(A.order.find((id) => A.P.items.get(id).pid === pid));
      const html = A.units.map((u) => u.html).join(' ');
      for (const id of A.order) if (A.P.items.get(id).pid === pid && !html.includes(`data-id="${id}"`)) miss.push(id);
      for (const e of document.querySelectorAll('#text .pn.ins')) { checked++; if (e.getBoundingClientRect().width < 9) narrow++; }
    }
    return { miss, narrow, checked, dir: getComputedStyle(document.getElementById('text')).direction, hdir: document.documentElement.dir };
  });
  ok(!r.miss.length, `${file}: פריטים בלי סימון בטקסט: ${r.miss.join(',')}`);
  ok(!r.narrow, `${file}: ${r.narrow} מתוך ${r.checked} סימני פיסוק צרים מ-9 פיקסלים`);
  ok(r.dir === 'rtl' && r.hdir === 'rtl', `${file}: כיוון ${r.dir}/${r.hdir}`);
  await p.context().close();
}

async function screens(file) {
  const p = await open(file, SIZES[0]);
  for (const sc of ['list', 'more', 'sum']) {
    await p.evaluate((sc) => __pane.screen(sc), sc);
    ok(await p.evaluate(() => !document.getElementById('screen').hidden && document.documentElement.scrollWidth <= innerWidth), `${file}: מסך ${sc}`);
  }
  const dl = p.waitForEvent('download');
  await p.click('#exp');
  const d = await dl; const data = JSON.parse(fs.readFileSync(await d.path(), 'utf8'));
  ok(data.items && data.items.length && data.siman, `${file}: ייצוא החלטות`);
  await p.context().close();
}

async function shots(dir) {
  fs.mkdirSync(dir, { recursive: true });
  const KH = pkgFile('קה'), KD = pkgFile('קד');
  for (const sz of SIZES) {
    const tag = sz.join('x');
    const p = await open(KH, sz, 2);
    const shot = (name) => p.screenshot({ path: path.join(dir, `${tag} - ${name}.png`) });
    const go = async (id) => { await p.evaluate((id) => __pane.go(id), id); await p.waitForTimeout(1300); };
    // לרצף: פריט פתוח ששני הפריטים הפתוחים אחריו באותו קטע
    const seqId = await p.evaluate(() => { const A = __pane, o = A.order.filter((id) => A.S.state(id).state === 'open' && A.P.items.get(id).kind === 'punct');
      for (let i = 30; i < o.length - 2; i++) { const it = o.slice(i, i + 3).map((x) => A.P.items.get(x)); if (it.every((x) => x.pid === it[0].pid)) { A.go(o[i]); const [a, b] = A.chunks[A.chunk]; if (o.slice(i, i + 3).every((x) => A.unitOf.get(x) >= a && A.unitOf.get(x) <= b) && A.order.indexOf(o[i + 2]) - A.order.indexOf(o[i]) === 2) return o[i]; } } return null; });
    await go('i0002'); await shot('01 פסיק');
    await go('i0042'); await shot('02 שינוי מילה');
    await go('i0007'); await shot('03 ראשי תיבות');
    await go('i0057'); await shot('04 מראה מקום עם ציטוט');
    await go('i0071'); await shot('05 מראה מקום בלי ציטוט');
    await go('i0068'); await shot('06 כותרת צד');
    await go('i0112'); await shot('07 לבירור עם הסבר ארוך');
    await go('i0042'); await p.click('#bOt'); await shot('08 שדה אחר פתוח');
    await p.click('#othNo');
    // רצף: שתי החלטות באותו קטע
    await go(seqId); await shot('09 רצף א - לפני');
    await p.click('#bOk'); await p.waitForTimeout(150); await shot('10 רצף ב - אחרי אישור');
    await p.click('#bNo'); await p.waitForTimeout(150); await shot('11 רצף ג - אחרי דחייה');
    await p.evaluate(() => __pane.screen('list')); await shot('12 מסך רשימה');
    await p.evaluate(() => __pane.screen('more')); await shot('13 מסך עוד');
    await p.evaluate(() => __pane.screen('sum')); await shot('14 מסך סיכום');
    await p.context().close();
    const q = await open(KD, sz, 2);
    await q.evaluate(() => __pane.go('i0008')); await q.waitForTimeout(800); await q.screenshot({ path: path.join(dir, `${tag} - 15 סימן שאלה כפול.png`) });
    await q.context().close();
  }
}

(async () => {
  browser = await chromium.launch();
  const t0 = Date.now();
  for (const f of PKGS) await marks(f);
  console.log('סימונים וכיוון: נבדק');
  for (const s of ['קה', 'קד']) console.log(`יציבות ${s}: ${await stability(pkgFile(s))} מקרים נמדדו באותו קטע`);
  for (const s of ['קה', 'קו']) console.log(`הליכה ${s}: ${await walk(pkgFile(s), 140)}`);
  await screens(pkgFile('קה'));
  if (SHOTS) { await shots(SHOTS); console.log('צילומים: ' + SHOTS); }
  await browser.close();
  console.log(`${pass} עברו, ${fail} נכשלו (${Math.round((Date.now() - t0) / 1000)} שניות)`);
  process.exit(fail ? 1 : 0);
})();
