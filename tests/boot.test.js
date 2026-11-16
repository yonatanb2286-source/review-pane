/* node tests/boot.test.js <חבילה.json בנתיב בלי עברית> [צילום.png] — החלונית בתוך "Word" מדומה: מסך טעינה, בחירת קובץ, 3 אישורים, פתיחה מחדש משחזרת מ-IndexedDB. */
// בדיקת אתחול בתוך "Word" מדומה: מסך טעינה, בחירת קובץ, אישור, ושמירה ב-IndexedDB אחרי טעינה מחדש
const { chromium } = require('playwright'); const fs = require('fs'), path = require('path');
const ROOT = process.cwd(), PK = process.argv[2], SHOT = process.argv[3];
const mock = fs.readFileSync(path.join(ROOT, 'tests/wordMock.js'), 'utf8');
const pkg = JSON.parse(fs.readFileSync(PK, 'utf8'));
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 329, height: 600 } }); const p = await ctx.newPage();
  const errs = []; p.on('pageerror', (e) => errs.push(e.message));
  const officeJs = `(function(){var module={exports:{}};var global=window;${mock};var M=module.exports;
    var d=new M.Doc(); d.addPara('סימן קודם','1_סימן');
    (${JSON.stringify(pkg.paragraphs.map((x) => [x.orig.trim(), x.style]))}).forEach(function(a){d.addPara(a[0],a[1]);});
    d.addPara('סימן אחרי','1_סימן');
    var saved=null; try{saved=JSON.parse(sessionStorage.getItem('mockdoc'));}catch(e){}
    if(saved){d.paras=saved;}
    window.__doc=d; M.install(d);
    window.Office={HostType:{Word:'Word'},onReady:function(cb){setTimeout(function(){cb({host:'Word'})},30)},context:{document:{addHandlerAsync:function(){}}},EventType:{DocumentSelectionChanged:'s'}};
  })();`;
  await p.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'application/javascript', body: officeJs }));
  await p.goto('file://' + ROOT + '/review.html'); await p.waitForTimeout(800);
  const t1 = await p.evaluate(() => document.getElementById('screen').innerText);
  console.log('1. מסך פתיחה:', t1.includes('טעינת סימן') ? 'מסך טעינה ✓' : 'חסר! ' + t1.slice(0, 80));
  console.log('pick:', await p.evaluate(() => { const e = document.getElementById('pick'); return !!e + ' ' + (e && typeof e.onchange); })); await p.setInputFiles('#pick', PK); await p.waitForTimeout(1500); console.log('msg:', await p.evaluate(() => (document.getElementById('loadMsg')||{}).innerText));
  const st = await p.evaluate(() => document.getElementById('sTitle').innerText + ' | ' + document.getElementById('sTrack').innerText);
  console.log('2. אחרי טעינה:', st, '/', await p.evaluate(() => document.getElementById('screen').innerText.slice(0,200)));
  if (SHOT) await p.screenshot({ path: SHOT });
  // אישור 3 פריטים במקלדת/כפתור
  for (let k = 0; k < 3; k++) { await p.evaluate(() => { const b = [...document.querySelectorAll('#row1 button')].find((x) => /אישור|אשר/.test(x.innerText)); b && b.click(); }); await p.waitForTimeout(400); }
  const w = await p.evaluate(() => ({ ops: window.__doc.ops, mine: window.__doc.paras.flatMap((x) => x.chars).filter((c) => c.add === 'y.m. b').length }));
  console.log('3. כתיבות במסמך:', JSON.stringify(w));
  await p.waitForTimeout(2500);
  await p.evaluate(() => sessionStorage.setItem('mockdoc', JSON.stringify(window.__doc.paras)));
  await p.reload(); await p.waitForTimeout(1500);
  const st2 = await p.evaluate(() => document.getElementById('sTitle').innerText + ' | ' + document.getElementById('sLeft').innerText + ' | screen hidden=' + document.getElementById('screen').hidden);
  console.log('4. אחרי פתיחה מחדש (IndexedDB):', st2);
  console.log('שגיאות:', errs.length ? errs : 'אין');
  await b.close();
})();
