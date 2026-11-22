/* review.js — חלונית הסקירה (שלב 3): מסך ראשי עם טקסט יציב, מסך "עוד", רשימה, סיכום, מצב הדגמה.
   מדבר עם המסמך רק דרך ReviewCore.Session ו-port. ב-Word ה-port הוא WordPortWord; מחוץ ל-Word (הדגמה ובדיקות) WordPortFake.
   כלל העיצוב: אחרי החלטה לא בונים מחדש את הטקסט, רק מחליפים מחלקות (צבע/קו), כך ששום מילה לא זזה. */
(function () {
  'use strict';
  const C = window.ReviewCore, $ = (id) => document.getElementById(id);
  const KIND = { punct: 'פיסוק', word: 'מילה', spelling: 'כתיב', abbr: 'ר"ת', ref: 'מראה מקום', heading: 'כותרת צד', split: 'פיצול', translation: 'תרגום', query: 'לבירור', info: 'לידיעה', topic: 'נושא', tech: 'טכני' };
  const STATE = { open: 'פתוח', accepted: 'אושר', rejected: 'נדחה', other: 'אחר', seen: 'נראה', info: 'לידיעה', changed: 'השתנה בקובץ', notfound: 'לא נמצא', na: 'לא רלוונטי', blocked: 'ממתין לפריט אחר' };
  const REF = { verified: 'אומת', corrected: 'תוקן', not_found: 'לא נמצא', from_history: 'מההיסטוריה', unchecked: 'לא נבדק' };
  const NEED = new Set(['open', 'blocked', 'changed', 'notfound', 'info']);
  const PUNCT = /^[\s.,:;!?״"'׳\-–—()\[\]]*$/;
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const hhmm = (iso) => { try { return new Date(iso).toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' }); } catch (e) { return ''; } };

  const A = { S: null, P: null, port: null, order: [], pos: new Map(), cur: null, pid: null, units: null, chunks: null, chunk: -1,
    mode: null, armed: null, times: {}, msg: null, filter: { st: 'need', kind: null }, hist: [], reports: [], repDone: new Set(), wideN: 2, settings: { hl: true, mark: true } };
  try { Object.assign(A.settings, JSON.parse(localStorage.getItem('review-pane-settings') || '{}')); } catch (e) { /* בלי הגדרות שמורות */ }
  window.__pane = A;   // לבדיקות

  // ───────── טעינה
  async function start(pkg, port, saved) {
    A.started = false;          // עד שהטעינה מצליחה לא שומרים דבר (כדי שחבילה שנכשלה לא תחליף את הקודמת)
    A.port = port || (A.inWord ? new WordPortWord.WordPort() : WordPortFake.FakeDoc.fromPackage(pkg));
    A.times = (saved && saved.times) || {};
    A.reports = A.inWord ? await allReports((saved && saved.reports) || []) : ((saved && saved.reports) || []);
    A.port.highlightHeadings = A.settings.hl;
    A.S = new C.Session(A.port, pkg, { log: (saved && saved.log) || {} });
    A.P = A.S.P;
    if (A.port.setPackage) A.port.setPackage(A.P);
    if (A.port.onCursorMoved) A.port.onCursorMoved(onCursor);
    await A.S.refresh();
    const pIx = new Map(pkg.paragraphs.map((p, i) => [p.pid, i]));
    const posOf = (it) => (['heading_before', 'topic', 'para_before', 'split_before'].includes(it.op) ? it.from - 0.5 : it.from) + (it.rel_start || 0) / 1e4 + (it.order || 0) / 1e6;
    A.order = pkg.items.filter((it) => it.kind !== 'tech')
      .sort((a, b) => (pIx.get(a.pid) - pIx.get(b.pid)) || (posOf(a) - posOf(b)) || (a.id < b.id ? -1 : 1)).map((it) => it.id);
    A.order.forEach((id, i) => A.pos.set(id, i));
    $('ver').textContent = 'גרסה ' + (window.PANE_VERSION || '');
    await track();
    const first = A.order.find((id) => NEED.has(st(id))) || A.order[0];
    if (first) go(first); else empty('אין פריטים בחבילה');
    window.addEventListener('resize', () => { if (A.pid) { const c = A.cur; A.pid = null; go(c); } });
    A.started = true;
    if (A.inWord) { clearInterval(A.saveT); A.saveT = setInterval(save, 2000); save(); }
  }
  /* הדיווחים משותפים לכל החבילות: נאספים מכל מה שנשמר אי פעם (גם מגרסה 0.5 ששמרה אותם לפי חבילה) */
  const repKey = (r) => r.time + '|' + r.note;
  const newReps = () => A.reports.filter((r) => !(A.repDone || new Set()).has(repKey(r)));
  async function allReports(extra) {
    try {
      const ex = await DB.get('reportsExported');
      A.repDone = new Set(ex || []);
      A.repSeed = !ex;     // לפני 0.5.3 לא נרשם מה יוצא: כל מה שנכתב עד הייצוא של 22.11 כבר נשלח
    } catch (e) { A.repDone = new Set(); }
    let all = [];
    try { all = (await DB.get('reports')) || []; } catch (e) { /* */ }
    try { for (const v of await DB.values()) if (v && Array.isArray(v.reports)) all = all.concat(v.reports); } catch (e) { /* */ }
    all = all.concat(extra || []);
    if (A.repSeed) { all.forEach((r) => { if (r.time < '2026-11-22T13:30:00Z') A.repDone.add(r.time + '|' + r.note); }); A.repSeed = false; try { await DB.put('reportsExported', [...A.repDone]); } catch (e) { /* */ } }
    const seen = new Set(), out = [];
    for (const r of all) { const k = r.time + '|' + r.note; if (!seen.has(k)) { seen.add(k); out.push(r); } }
    out.sort((a, b) => (a.time < b.time ? -1 : 1));
    return out;
  }
  // ───────── שמירה ב-IndexedDB (החבילה + היומן), כדי שבפתיחה הבאה של Word העבודה תימשך
  const DB = {
    open() { return new Promise((ok, no) => { try { const r = indexedDB.open('review-pane', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); } catch (e) { no(e); } }); },
    async get(k) { const db = await DB.open(); return new Promise((ok, no) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => ok(q.result); q.onerror = () => no(q.error); }); },
    async values() { const db = await DB.open(); return new Promise((ok, no) => { const q = db.transaction('kv').objectStore('kv').getAll(); q.onsuccess = () => ok(q.result || []); q.onerror = () => no(q.error); }); },
    async put(k, v) { const db = await DB.open(); return new Promise((ok, no) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => ok(); t.onerror = () => no(t.error); }); },
  };
  const pkgKey = (pk) => 'log:' + pk.siman + ':' + pk.round + ':' + (pk.built || '');
  async function save() {
    if (!A.S || !A.started) return;
    const pk = A.P.pkg, rec = { log: A.S.log, times: A.times };
    const j = JSON.stringify(rec) + A.reports.length; if (j === A.savedJ) return;
    try { await DB.put(pkgKey(pk), rec); await DB.put('reports', A.reports); if (!A.savedPkg) { await DB.put('cur', pk); A.savedPkg = true; } A.savedJ = j; } catch (e) { A.saveErr = e.message; }
  }
  async function loadPkg(pkg) {
    // חבילה חדשה מחליפה את הנוכחית; היומן של כל חבילה נשמר בנפרד ולא נמחק
    if (A.S) await save();
    A.savedPkg = false; A.savedJ = null;
    let saved = null; if (A.inWord) { try { saved = await DB.get(pkgKey(pkg)); } catch (e) { /* בלי שמירה */ } }
    A.pid = null;
    await start(pkg, null, saved);
  }
  const st = (id) => { const s = A.S.state(id); return s ? s.state : 'notfound'; };
  const item = (id) => A.P.items.get(id);
  const left = () => A.order.filter((id) => NEED.has(st(id))).length;

  async function track() {
    let on = await A.port.getTracking();
    if (!on && A.port.setTrackingOn) { await A.port.setTrackingOn(); on = await A.port.getTracking(); }
    const el = $('sTrack'); el.textContent = on ? 'מעקב ✓' : 'מעקב כבוי!'; el.title = on ? 'מעקב השינויים פעיל' : 'מעקב השינויים כבוי: לא נכתב דבר'; el.className = 'chip' + (on ? '' : ' off');
    A.trackOn = on;
  }

  // ───────── בניית הטקסט היציב של פסקה: יחידות (תיבה או אשכול), כל אחת עם הכותרות/פיצול שלפניה
  function buildUnits(par) {
    const P = A.P, units = [], unitOf = new Map();
    const its = par.items.map((i) => P.items.get(i)).filter((it) => it.kind !== 'tech');
    const before = new Map(), marks = new Map();
    for (const it of its) {
      if (it.cluster) continue;
      if (it.op === 'none') { for (let t = it.from; t <= Math.max(it.from, it.to); t++) { if (!marks.has(t)) marks.set(t, []); marks.get(t).push(it); } continue; }
      if (!before.has(it.from)) before.set(it.from, []);
      before.get(it.from).push(it);
    }
    const cls = new Map((P.clustersOf.get(par.pid) || []).map((c) => [c.from, c]));
    const n = par.orig_tokens.length;
    for (let t = 0; t < n;) {
      let h = '', ids = [];
      const pre = (before.get(t) || []).slice().sort((a, b) => (a.op === 'split_before' ? 1 : 0) - (b.op === 'split_before' ? 1 : 0) || (a.order || 0) - (b.order || 0));
      const splits = pre.filter((x) => x.op === 'split_before'), heads = pre.filter((x) => x.op !== 'split_before');
      // פיצול קודם לכותרת שנכנסת במקום הפיצול
      for (const s of splits) { h += `<span class="it k-split sp" data-id="${s.id}">¶</span>`; ids.push(s.id); }
      for (const x of heads) { h += `<span class="hd it ins k-${x.kind}" data-id="${x.id}">${esc(x.new)}</span>`; ids.push(x.id); }
      const c = cls.get(t);
      let end;
      if (c) {
        const items = c.items.map((iid, k) => ({ it: P.items.get(iid), k })).sort((a, b) => (a.it.rel_start - b.it.rel_start) || (a.k - b.k));
        let p = 0, body = '';
        const plain = (s) => s.split(/(\s+)/).map((w) => (/^\s+$/.test(w) ? ' ' : w ? `<span class="w">${esc(w)}</span>` : '')).join('');
        for (const { it } of items) {
          if (it.rel_start > p) body += plain(c.orig.slice(p, it.rel_start));
          const old = c.orig.slice(it.rel_start, it.rel_start + it.char_len);
          const pn = (s) => (PUNCT.test(s) && s.trim().length <= 2 ? ' pn' : '');
          const qf = (it.flags || []).includes('??') ? ' qflag' : '';
          if (old) body += `<span class="it del k-${it.kind}${pn(old)}" data-id="${it.id}">${esc(old)}</span>`;
          if (it.new) {
            const lead = /^\s/.test(it.new) ? ' ' : '', tail = /\s$/.test(it.new) && it.new.trim() ? ' ' : '';
            if (it.new.trim()) body += lead + `<span class="it ins k-${it.kind}${pn(it.new.trim())}${qf}" data-id="${it.id}">${esc(it.new.trim())}</span>` + tail;
          }
          ids.push(it.id);
          p = it.rel_start + it.char_len;
        }
        if (p < c.orig.length) body += plain(c.orig.slice(p));
        const mk = []; for (let x = c.from; x <= c.to; x++) (marks.get(x) || []).forEach((m) => mk.includes(m) || mk.push(m));
        mk.forEach((m) => ids.push(m.id));
        let w = `<span class="cl">${body}</span>`;
        for (const m of mk.slice().reverse()) w = `<span class="it k-${m.kind}" data-id="${m.id}">${w}</span>`;
        h += w;
        end = c.orig; t = c.to + 1;
      } else {
        const mk = marks.get(t) || [];
        let w = `<span class="w">${esc(par.orig_tokens[t])}</span>`;
        for (const m of mk.slice().reverse()) { w = `<span class="it k-${m.kind}" data-id="${m.id}">${w}</span>`; ids.push(m.id); }
        h += w;
        end = par.orig_tokens[t]; t++;
      }
      // מידע שמסומן על תיבה שבתוך אשכול
      ids.forEach((id) => unitOf.set(id, units.length));
      units.push({ html: `<span class="u" data-u="${units.length}">${h}</span>`, end });
    }
    for (const it of its) if (!unitOf.has(it.id)) unitOf.set(it.id, Math.min(units.length - 1, Math.max(0, it.from)));
    return { units, unitOf };
  }

  // חלוקה קבועה לקטעים: כל קטע ממלא את האזור בלי גלילה; חיתוך בסוף משפט, אחרת אחרי פסיק
  function chunkUnits(units) {
    const box = $('text'), M = $('measure');
    M.style.width = box.offsetWidth + 'px';
    const cs = getComputedStyle(box);
    const H = box.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) - 1;
    const out = [];
    for (let i = 0; i < units.length;) {
      M.innerHTML = units.slice(i, i + 400).map((u) => u.html).join(' ');
      const top = M.getBoundingClientRect().top + parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth);
      const els = M.children;
      let last = -1;
      for (let k = 0; k < els.length; k++) {
        const r = els[k].getClientRects(); const b = r.length ? r[r.length - 1].bottom : 0;
        if (b - top <= H) last = k; else break;
      }
      if (last < 0) last = 0;
      let cut = last;
      if (i + last + 1 < units.length) {
        const lo = Math.floor(last * 0.55);
        let k = last; for (; k >= lo; k--) if (/[.:?!]["'״׳)]*$/.test(units[i + k].end)) break;
        if (k < lo) { for (k = last; k >= lo; k--) if (/[,;]["'״׳)]*$/.test(units[i + k].end)) break; }
        if (k >= lo) cut = k;
      }
      out.push([i, i + cut]); i += cut + 1;
    }
    M.innerHTML = '';
    return out;
  }

  // ───────── מעבר לפריט
  function go(id, fromCursor) {
    const it = item(id); if (!it) return;
    A.mode = null; A.msg = null; A.wideN = 2;
    if (A.pid !== it.pid) {
      A.pid = it.pid;
      const b = buildUnits(A.P.paras.get(it.pid));
      A.units = b.units; A.unitOf = b.unitOf; A.chunks = chunkUnits(A.units); A.chunk = -1;
    }
    const u = A.unitOf.get(id) || 0;
    const ch = A.chunks.findIndex(([a, b]) => u >= a && u <= b);
    if (ch !== A.chunk) {
      const turned = A.chunk !== -1 || A.flashNext;
      if (ch < 0) return;
      A.chunk = ch;
      const [a, b] = A.chunks[ch];
      $('text').innerHTML = A.units.slice(a, b + 1).map((x) => x.html).join(' ');
      if (turned) flash();
    }
    A.flashNext = false;
    A.cur = id; A.reading = false;
    paintStates();
    renderStatus(); renderDet(); renderBtns();
    if (!fromCursor) selectInWord(id);
  }
  // קריאה ברצף: דפדוף קטע אחר קטע בכל הסימן, גם בפסקאות בלי שינויים
  function showChunk(pid, ch) {
    if (A.pid !== pid) { A.pid = pid; const b = buildUnits(A.P.paras.get(pid)); A.units = b.units; A.unitOf = b.unitOf; A.chunks = chunkUnits(A.units); }
    if (ch < 0) ch = A.chunks.length - 1;
    const [a, b] = A.chunks[ch];
    const here = A.order.filter((id) => item(id).pid === pid && (A.unitOf.get(id) || 0) >= a && (A.unitOf.get(id) || 0) <= b);
    const pick = here.find((id) => NEED.has(st(id))) || here[0];
    A.chunk = -2;          // מכריח ציור מחדש
    if (pick) { A.reading = false; go(pick); A.reading = false; return; }
    A.chunk = ch; A.reading = true; A.mode = null; A.msg = null;
    $('text').innerHTML = A.units.slice(a, b + 1).map((x) => x.html).join(' ');
    flash(); paintStates(); renderStatus(); renderDet(); renderBtns();
    if (A.port.selectParagraph) { const m = A.S.S.match.map.get(pid); if (m && m.keys && m.keys[0]) { A.selfMove = Date.now(); A.port.selectParagraph(m.keys[0]).catch(() => {}); } }
  }
  function page(dir) {
    const paras = A.P.pkg.paragraphs, pi = paras.findIndex((p) => p.pid === A.pid);
    const ch = A.chunk + dir;
    if (ch >= 0 && ch < A.chunks.length) return showChunk(A.pid, ch);
    const np = paras[pi + dir]; if (!np) return;
    showChunk(np.pid, dir > 0 ? 0 : -1);
  }
  function shrinkImg(file) {   // תמונה לדיווח: עד 1400 פיקסלים, JPEG
    return new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => { const im = new Image(); im.onload = () => { const k = Math.min(1, 1400 / Math.max(im.width, im.height)); const c = document.createElement('canvas'); c.width = Math.round(im.width * k); c.height = Math.round(im.height * k); c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); res(c.toDataURL('image/jpeg', 0.8)); }; im.onerror = rej; im.src = fr.result; };
      fr.onerror = rej; fr.readAsDataURL(file);
    });
  }
  function flash() { const t = $('text'), p = $('sPos'); t.classList.add('turn'); p.classList.add('flash'); setTimeout(() => { t.classList.remove('turn'); p.classList.remove('flash'); }, 700); }
  function paintStates() {
    for (const el of $('text').querySelectorAll('.it[data-id]')) {
      const id = el.dataset.id;
      el.className = el.className.replace(/\s*(st-\S+|cur)/g, '') + ' st-' + st(id) + (id === A.cur ? ' cur' : '');
    }
  }
  async function selectInWord(id) {
    const s = A.S.state(id); if (!s || !s.loc || !A.port.select) return;
    A.selfMove = Date.now(); A.lastCur = s.loc.key;
    const from = s.loc.from !== undefined ? s.loc.from : s.loc.token, to = s.loc.from !== undefined ? Math.max(s.loc.from, s.loc.to) : s.loc.token;
    try { if (from !== undefined) await A.port.select(s.loc.key, from, to); else await A.port.selectParagraph(s.loc.key); } catch (e) { /* בחירה לא חיונית */ }
  }
  // תזוזת סמן ב-Word על ידי יונתן (לא כזו שהחלונית גרמה)
  async function onCursor() {
    // מגיבים רק כשיונתן עבר לפסקה אחרת ב-Word. כתיבה או בחירה של החלונית, או הקלדה באותה פסקה, לא מקפיצות את החלונית.
    if (A.busy || (A.selfMove && Date.now() - A.selfMove < 1500)) return;
    const key = await A.port.getCursorParagraph(); if (!key || key === A.lastCur) return;
    A.lastCur = key;
    const pid = [...A.S.S.match.map.entries()].find(([, m]) => m.keys && m.keys.includes(key));
    if (!pid) return;
    // באותה פסקה שבה עובדים: לא קופצים (פסקה שפוצלה היא כמה פסקאות ב-Word, וזה הקפיץ לפריטים שדילגת עליהם – קיב 22.11)
    if (pid[0] === A.pid) return;
    const ids = A.order.filter((id) => item(id).pid === pid[0]);
    const nid = ids.find((id) => NEED.has(st(id))) || ids[0];
    if (nid && nid !== A.cur) go(nid, true);
  }
  function step(dir, needOnly) {
    const n = A.order.length; let i = A.pos.get(A.cur);
    for (let k = 1; k <= n; k++) {
      const j = (i + dir * k + n) % n;
      if (!needOnly || NEED.has(st(A.order[j]))) return go(A.order[j]);
    }
    if (needOnly) { renderAll(); }
  }
  // אחרי החלטה: רק קדימה. אם לא נשאר פתוח בהמשך, לא קופצים להתחלה
  function advance() {
    if (!left()) return renderAll();
    const i = A.pos.get(A.cur);
    for (let j = i + 1; j < A.order.length; j++) if (NEED.has(st(A.order[j]))) return go(A.order[j]);
    A.msg = 'אין עוד פריטים פתוחים מכאן ועד סוף הסימן. נשארו ' + left() + ' פתוחים לפני כן ("הבא" ימשיך מההתחלה).'; renderAll();
  }
  // מעבר לפסקה שבה הסמן של Word עומד עכשיו
  async function toCursor() {
    const key = A.port.getCursorParagraph ? await A.port.getCursorParagraph() : null;
    const pid = key && [...A.S.S.match.map.entries()].find(([, m]) => m.keys && m.keys.includes(key));
    if (!pid) { A.msg = 'הסמן ב-Word אינו בפסקה של הסימן שנטען'; return renderDet(); }
    A.lastCur = key;
    const ids = A.order.filter((id) => item(id).pid === pid[0]);
    const nid = ids.find((id) => NEED.has(st(id))) || ids[0];
    if (nid) return go(nid, true);
    showChunk(pid[0], 0);
  }
  function renderAll() { paintStates(); renderStatus(); renderDet(); renderBtns(); }

  // ───────── שורת מצב
  function renderStatus() {
    const pk = A.P.pkg, pi = pk.paragraphs.findIndex((p) => p.pid === A.pid) + 1;
    const withItems = pk.paragraphs.filter((p) => p.items.length).length;
    $('sTitle').textContent = `סימן ${pk.siman} · סבב ${pk.round}`;
    $('sLeft').textContent = `נותרו ${left()} מתוך ${A.order.length}`;
    $('sPos').textContent = `פסקה ${pi}/${pk.paragraphs.length}` + (A.chunks.length > 1 ? ` · קטע ${A.chunk + 1}/${A.chunks.length}` : '') + (A.reading ? ' · קריאה' : '');
    $('sPos').title = `${withItems} פסקאות עם שינויים`;
  }

  // ───────── פרטי השינוי
  function beforeAfter(it) {
    const short = (s, n) => { s = String(s).trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
    if (it.cluster) {
      const c = A.P.clusters.get(it.cluster), k = c.items.indexOf(it.id);
      const o = c.orig, nw = C.outcomeText(c, A.P.items, 1 << k);
      if (it.kind === 'ref' && !it.char_len) return `<span class="n">+ ${esc(short(it.new, 44))}</span>`;
      // רק מה שמשתנה מסומן: חלק משותף רגיל, מחיקה בקו חוצה, הוספה בירוק (פסיק = רק הפסיק)
      let p = 0; while (p < o.length && p < nw.length && o[p] === nw[p]) p++;
      let q = 0; while (q < o.length - p && q < nw.length - p && o[o.length - 1 - q] === nw[nw.length - 1 - q]) q++;
      if (!(PUNCT.test(o.slice(p, o.length - q)) && PUNCT.test(nw.slice(p, nw.length - q)))) {   // שינוי באותיות: מסמנים מילה שלמה
        while (p > 0 && !/\s/.test(o[p - 1])) p--;
        while (q > 0 && !/\s/.test(o[o.length - q])) q--;
      }
      const pre = o.slice(0, p), del = o.slice(p, o.length - q), ins = nw.slice(p, nw.length - q), suf = o.slice(o.length - q);
      const cut = (x, n) => (x.length > n ? x.slice(0, n - 1) + '…' : x);
      return `${esc(cut(pre, 24))}${del ? `<span class="o">${esc(cut(del, 30))}</span>` : ''}${ins ? `<span class="n">${esc(cut(ins, 34))}</span>` : ''}${esc(cut(suf, 24))}`;
    }
    if (it.op === 'split_before') return `פסקה חדשה מ: <span class="n">«${esc(A.P.paras.get(it.pid).orig_tokens.slice(it.from, it.from + 3).join(' '))}»</span>`;
    if (['heading_before', 'topic', 'para_before'].includes(it.op)) return 'שורה חדשה (בראש הטקסט)';
    return '';
  }
  function quoteHtml(rc) {
    if (!rc) return '';
    let h = `<div><span class="rs rs-${rc.gray ? 'gray' : rc.status}">${rc.gray ? 'אומת בסבב קודם' : (REF[rc.status] || rc.status)}</span></div>`;
    const q = rc.quote;
    if (q && q.text) {
      const keys = (q.key || []).map(C.norm);
      const txt = q.text.split(/(\s+)/).map((w) => (/^\s+$/.test(w) ? w : keys.some((k) => { const n = C.norm(w); return k && (n === k || (n.length === k.length + 1 && n.endsWith(k) && 'והבכלמש'.includes(n[0]))); }) ? `<b>${esc(w)}</b>` : esc(w))).join('');
      h += `<div class="quote"><div class="qsrc">${esc(q.source || '')} · <span dir="ltr">${esc(q.ref || '')}</span>${q.from_book ? ' · מהספר, לא מהמקור' : ''}</div><div class="qtxt">${txt}</div></div>`;
    } else h += `<div class="quote none">אין ציטוט${rc.no_quote_reason ? ': ' + esc(rc.no_quote_reason) : ''}</div>`;
    return h;
  }
  function renderDet() {
    if (A.reading) { $('detIn').innerHTML = '<div class="note"><b>קריאה ברצף.</b> אין שינויים בקטע הזה.</div><div class="note hint">▼ או PageDown לקטע הבא, ▲ או PageUp לקודם. "הבא" עובר לשינוי הפתוח הבא.</div>'; moreHint(); return; }
    const it = item(A.cur), s = st(A.cur), box = $('detIn');
    let h = `<div class="dh"><span class="tag k-${it.kind}">${KIND[it.kind] || it.kind}</span><span class="ba">${beforeAfter(it)}</span><span class="stt">${STATE[s] || s}</span></div>`;
    if (A.msg) h += `<div class="err">${esc(A.msg)}</div>`;
    const qq = (it.flags || []).includes('??') || it.kind === 'query' || /\?\?/.test(it.new || '');
    if (A.mode === 'other') {
      const ss = A.S.state(A.cur) || {};
      let cur, lab = 'הנוסח שלך (ייכתב במעקב בשמך במקום ההצעה):';
      A.wide = null;
      if (it.op === 'none' && !ss.loc) cur = it.old;
      else if (it.op === 'none') {
        const lv = A.S.S.liveByKey.get(ss.loc.key).tokens, from = Math.max(0, ss.loc.from - A.wideN), to = Math.min(lv.length - 1, ss.loc.to + A.wideN);
        cur = lv.slice(from, to + 1).join(' '); A.wide = { key: ss.loc.key, from, to };
        lab = 'המקום עם המילים שסביבו. אפשר לתקן כאן, להוסיף ציון או ד"ה. רק מה שתשנה ייכתב במעקב בשמך:';
      } else if (it.cluster && ss.loc && ss.mask !== undefined) {
        // שדה רחב: ההצעה עם שתי מילים לפניה ואחריה, כדי שאפשר יהיה גם להזיז אותה
        const c = A.P.clusters.get(it.cluster), lv = A.S.S.liveByKey.get(ss.loc.key).tokens;
        const from = Math.max(0, ss.loc.from - A.wideN), to = Math.min(lv.length - 1, ss.loc.to + A.wideN);
        const mid = C.splitWords(C.outcomeText(c, A.P.items, ss.mask | (1 << c.items.indexOf(it.id))));
        cur = lv.slice(from, ss.loc.from).concat(mid, lv.slice(ss.loc.to + 1, to + 1)).join(' ');
        A.wide = { key: ss.loc.key, from, to };
        lab = 'ההצעה עם המילים שסביבה. אפשר לשנות כל דבר כאן, גם להזיז את הפסיק. רק מה שתשנה ייכתב במעקב בשמך:';
      } else cur = it.cluster ? it.new.trim() : it.new;
      const wk = A.wide ? A.wide.key + ':' + A.wide.from + ':' + A.wide.to : null;
      if (wk && wk !== A.wideSel && A.port.select) { A.wideSel = wk; A.selfMove = Date.now(); A.port.select(A.wide.key, A.wide.from, A.wide.to).catch(() => {}); }
      h += `<div class="oth"><label>${lab}</label><textarea id="othT" dir="rtl">${esc(cur)}</textarea>
        <label class="hint"><input type="checkbox" id="othMM" ${it.kind === 'ref' ? 'checked' : ''}> מה שבסוגריים (…) ייכתב בסגנון מ"מ</label>
        <div class="row"><button id="othOk" class="ok">כתוב</button><button id="othNo">ביטול</button>${A.wide ? `<button id="othWide" title="עוד מילים לפני ואחרי (מבטל מה שהקלדת בשדה)">הרחב את השדה</button>` : ''}</div></div>`;
    } else if (A.mode === 'note') {
      h += `<div class="rnote"><label class="hint">${it.op === 'none' ? 'מה החלטת (נכנס לקובץ ההחלטות):' : 'סיבת הדחייה (לא חובה):'}</label><textarea id="noteT" dir="rtl">${esc((A.S.log[A.cur] || {}).why || '')}</textarea>
        <div class="row"><button id="noteOk">${it.op === 'none' ? 'ראיתי, עם הערה' : 'דחה עם הערה'}</button><button id="noteNo">ביטול</button></div></div>`;
    } else {
      if (qq) h += `<div class="qqbar">?? לבירור</div>`;
      if (s === 'changed') h += `<div class="err">המקום נערך בקובץ: הנוסח שם אינו המקור ואינו ההצעה. "אחר" ירשום את מה שכתוב שם.</div>`;
      if (s === 'notfound') h += `<div class="err">המקום לא נמצא בוודאות בקובץ, ולכן לא ייכתב. ${esc((A.S.state(A.cur) || {}).why || '')}</div>`;
      if (s === 'blocked') h += `<div class="err">צריך לאשר קודם את הפיצול שהכותרת תלויה בו.</div>`;
      // it.note הוא חיבור כל ההערות: כשיש הערות נפרדות מציגים רק אותן, כדי שלא ייכתב פעמיים (קיב 22.11)
      const notes = ((it.notes || []).length ? it.notes.map((x) => x.text) : [it.note]).filter((x, i, a) => x && a.indexOf(x) === i);
      if (it.ref_check && it.ref_check.text && !notes.includes(it.ref_check.text)) notes.push(it.ref_check.text);
      if (['heading_before', 'topic'].includes(it.op)) h += `<div class="note"><b>${esc(it.new)}</b></div>`;
      notes.forEach((x) => { h += `<div class="note">${esc(x)}</div>`; });
      if (it.kind === 'ref') h += quoteHtml(it.ref_check);
      const lg = A.S.log[A.cur];
      if (lg && lg.state === 'other' && lg.text) h += `<div class="note">הנוסח שלך: <b>${esc(lg.text)}</b></div>`;
      if (lg && lg.why) h += `<div class="note hint">הערתך: ${esc(lg.why)}</div>`;
    }
    box.innerHTML = h; box.scrollTop = 0;
    const t = $('othT') || $('noteT'); if (t) { t.focus(); t.setSelectionRange(t.value.length, t.value.length); }
    if ($('othWide')) $('othWide').onclick = () => { A.wideN = Math.min(30, A.wideN + 5); renderDet(); };
    if ($('othOk')) { $('othOk').onclick = () => doOther($('othT').value); $('othNo').onclick = () => { A.mode = null; renderDet(); renderBtns(); }; }
    if ($('noteOk')) { $('noteOk').onclick = () => (it.op === 'none' ? doSeen($('noteT').value) : doReject($('noteT').value)); $('noteNo').onclick = () => { A.mode = null; renderDet(); }; }
    moreHint();
  }
  function moreHint() { const b = $('detIn'); $('det').classList.toggle('more', b.scrollHeight - b.scrollTop - b.clientHeight > 4); }

  // ───────── כפתורים
  function renderBtns() {
    if (A.reading) { $('row1').innerHTML = '<button class="ok" disabled>אישור</button><button class="no" disabled>דחייה</button><button class="ot" disabled>אחר</button>'; $('bNote').disabled = true; return; }
    const s = st(A.cur), it = item(A.cur), r = $('row1');
    const lg = A.S.log[A.cur] || {};
    const isRef = it.kind === 'query' || it.kind === 'ref' || !!(it.ref_check && it.ref_check.text);
    if (it.op === 'none' && (s === 'info' || s === 'open')) r.innerHTML = `<button class="seen" id="bSeen">${isRef ? 'המקור תקין' : 'ראיתי'}</button><button class="ot" id="bOt"${(A.S.state(A.cur) || {}).loc && A.trackOn ? '' : ' disabled'}>אחר</button>`;
    else if (['accepted', 'rejected', 'other', 'seen', 'na'].includes(s)) {
      const tm = A.times[A.cur] || lg.time;
      r.innerHTML = `<div class="done ${s}">${s === 'seen' && lg.verdict ? 'המקור תקין' : STATE[s]}${tm ? ' · ' + hhmm(tm) : ''}</div><button class="reo" id="bReo">פתח מחדש</button>`;
    } else {
      const canOk = s === 'open';
      r.innerHTML = `<button class="ok" id="bOk"${canOk && A.trackOn ? '' : ' disabled'}>אישור</button><button class="no" id="bNo">דחייה</button><button class="ot" id="bOt"${(s === 'open' || s === 'changed') && A.trackOn ? '' : ' disabled'}>אחר</button>`;
    }
    $('bNote').disabled = !(s === 'open' || s === 'changed' || s === 'rejected' || (it.op === 'none' && ['info', 'seen', 'open'].includes(s)));
    $('bNote').textContent = it.op === 'none' ? 'הערה' : 'הערה לדחייה';
    if ($('bOk')) $('bOk').onclick = doApprove;
    if ($('bNo')) $('bNo').onclick = () => doReject('');
    if ($('bOt')) $('bOt').onclick = openOther;
    if ($('bSeen')) $('bSeen').onclick = () => doSeen();
    if ($('bReo')) $('bReo').onclick = doReopen;
  }

  // ───────── פעולות
  async function busy(fn) { if (A.busy) return; A.busy = true; try { await fn(); } catch (e) { A.msg = 'שגיאה: ' + e.message; renderDet(); } finally { A.busy = false; A.selfMove = Date.now(); } }
  function hist(id, act, r) { A.hist.push({ t: new Date().toISOString(), id, act, ok: !r || !!r.ok, why: r && !r.ok ? r.why : undefined, edit: r && r.edit ? r.edit.kind : undefined }); if (A.hist.length > 30) A.hist.shift(); }
  const doApprove = () => busy(async () => {
    const id = A.cur, r = await A.S.approve(id); hist(id, 'אישור', r);
    if (!r.ok) { A.msg = r.why; return renderAll(); }
    A.times[id] = new Date().toISOString(); advance(); warnAfter();
  });
  const doReject = (why) => busy(async () => {
    const id = A.cur; A.mode = null;
    if (!['open', 'changed', 'blocked', 'notfound', 'rejected'].includes(st(id))) return;
    await A.S.reject(id, why || ''); hist(id, 'דחייה'); A.times[id] = new Date().toISOString(); advance();
  });
  const doSeen = (why) => busy(async () => { const id = A.cur, it = item(id); A.mode = null; await A.S.markSeen(id);
    if (it.kind === 'query' || it.kind === 'ref' || (it.ref_check && it.ref_check.text)) A.S.log[id].verdict = 'ok'; if (typeof why === 'string' && why.trim()) A.S.log[id].why = why.trim(); A.times[id] = new Date().toISOString(); advance(); });
  function openOther() {
    if (st(A.cur) === 'changed') return busy(async () => { const id = A.cur; const r = await A.S.recordOther(id); if (!r.ok) { A.msg = r.why; return renderAll(); } A.times[id] = new Date().toISOString(); advance(); });
    A.mode = 'other'; renderDet();
  }
  const doOther = (text) => busy(async () => {
    const id = A.cur; text = String(text).trim();
    const it = item(id);
    const val = it.cluster && /^\s/.test(it.new) ? ' ' + text : text;
    if (!text) { A.msg = 'הנוסח ריק'; return renderDet(); }
    A.port.mmParens = !!($('othMM') && $('othMM').checked);
    const r = A.wide && (it.cluster || it.op === 'none') ? await A.S.otherRange(id, A.wide.key, A.wide.from, A.wide.to, text) : await A.S.other(id, val);
    A.port.mmParens = false; hist(id, 'אחר', r);
    A.mode = null;
    if (!r.ok) { A.msg = r.why; return renderAll(); }
    A.times[id] = new Date().toISOString(); advance(); warnAfter();
  });
  function warnAfter() { if (A.port.lastWarn) { A.msg = '⚠ בשינוי שנכתב עכשיו: ' + A.port.lastWarn; A.port.lastWarn = null; renderDet(); } }
  const doReopen = () => busy(async () => {
    const id = A.cur, s = st(id);
    if (['rejected', 'seen'].includes(s)) { delete A.S.log[id]; delete A.times[id]; await A.S.refresh(); renderAll(); }
    else { A.msg = 'השינוי הזה כבר נכתב בקובץ' + (A.times[id] ? ' (' + hhmm(A.times[id]) + ')' : '') + ', ולכן החלונית לא דוחה אותו. כדי לבטל: ב-Word לחיצה ימנית על השינוי ואז "דחה" (או Ctrl+Z מיד אחרי האישור). אחרי זה הפריט יחזור כאן ל"פתוח".'; renderDet(); }
  });
  async function bulk(kindOnly) {
    const ids = A.order.filter((id) => item(id).pid === A.pid && st(id) === 'open' && item(id).op !== 'none' && (!kindOnly || item(id).kind === 'punct'));
    let ok = 0, bad = 0;
    for (const id of ids) { const r = await A.S.approve(id); if (r.ok) { ok++; A.times[id] = new Date().toISOString(); } else bad++; }
    return { ok, bad };
  }

  // ───────── קלט
  $('text').addEventListener('click', (e) => { const el = e.target.closest('.it[data-id]'); if (el) go(el.dataset.id); });
  $('detIn').addEventListener('scroll', moreHint);
  $('bPrev').onclick = () => step(-1);
  $('bNext').onclick = () => step(1);
  $('bNote').onclick = () => { A.mode = 'note'; renderDet(); };
  $('bList').onclick = () => screen('list');
  $('bCur').onclick = () => toCursor();
  $('bPgUp').onclick = () => page(-1);
  $('bPgDn').onclick = () => page(1);
  $('bMore').onclick = () => screen('more');
  document.addEventListener('keydown', (e) => {
    if (!A.S) return;
    if (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'INPUT') {
      if (e.key === 'Escape') { A.mode = null; renderDet(); }
      else if (e.key === 'Enter' && !e.shiftKey && e.target.id === 'othT') { e.preventDefault(); doOther(e.target.value); }
      return;
    }
    if (!$('screen').hidden) { if (e.key === 'Escape') closeScreen(); return; }
    const k = e.key;
    if (k === 'Enter') { e.preventDefault(); if ($('bOk') && !$('bOk').disabled) doApprove(); else if ($('bSeen')) doSeen(); }
    else if (k === 'PageDown') { e.preventDefault(); page(1); }
    else if (k === 'PageUp') { e.preventDefault(); page(-1); }
    else if (k === 'Delete') { e.preventDefault(); if ($('bNo')) doReject(''); }
    else if (k === 'F2' || k === 'א') { e.preventDefault(); if ($('bOt') && !$('bOt').disabled) openOther(); }
    else if (k === 'ArrowLeft' || k === 'ArrowDown') { e.preventDefault(); step(1); }
    else if (k === 'ArrowRight' || k === 'ArrowUp') { e.preventDefault(); step(-1); }
  });

  // ───────── מסכים משניים
  function closeScreen() { $('screen').hidden = true; $('screen').innerHTML = ''; if (A.cur) { const c = A.cur; A.pid = null; go(c); } }
  function screen(name, arg) {
    const sc = $('screen'); sc.hidden = false;
    const back = `<button class="mini" id="scBack">חזרה ✕</button>`;
    let h = '';
    if (!A.P) name = 'load';                      // אין עדיין סימן טעון
    const pk = A.P ? A.P.pkg : null;
    if (name === 'more') {
      const here = A.order.filter((id) => item(id).pid === A.pid && st(id) === 'open' && item(id).op !== 'none');
      const np = here.filter((id) => item(id).kind === 'punct').length;
      const tech = pk.items.filter((it) => it.kind === 'tech');
      const techHere = tech.filter((it) => it.pid === A.pid);
      h += `<h2>עוד ${back}</h2>
        <div class="box"><b>דיווח על בעיה או בקשה</b><div class="hint">נרשם יחד עם המקום שבו אתה עומד ועם הפעולות האחרונות של החלונית.</div>
          <textarea id="repT" dir="rtl" style="width:100%;height:60px" placeholder="מה קרה, או מה היית רוצה"></textarea>
          <div class="row"><label class="hint" style="cursor:pointer;text-decoration:underline"><input type="file" id="repImg" accept="image/*" multiple hidden>צירוף תמונה</label> <span class="hint">(או הדבקה בתיבה עם Ctrl+V)</span> <span id="repImgN" class="hint">${A.repImgs && A.repImgs.length ? A.repImgs.length + ' תמונות מצורפות' : ''}</span></div>
          <div class="row"><button id="repOk">שמור דיווח</button><button id="repExp" ${newReps().length ? '' : 'disabled'}>ייצוא דיווחים חדשים (${newReps().length})</button></div><div id="repMsg" class="hint"></div></div>
        <div class="box"><b>הגדרות</b>
          <label class="hint" style="display:block"><input type="checkbox" id="setHl" ${A.settings.hl ? 'checked' : ''}> הזרחה צהובה לכותרת צד ול-(??) שנכתבים</label></div>
        ${A.port && A.port.diag ? `<div class="box"><b>דוח אבחון</b> <span class="hint">(להעתקה לשיחה אם משהו לא עבד)</span><textarea readonly style="width:100%;height:90px" dir="rtl">${esc(A.port.diag.join('\n'))}</textarea></div>` : ''}
        <div class="box"><b>אישור מרוכז בפסקה הנוכחית</b><div class="hint">כל שינוי נכתב בנפרד במעקב. לחיצה שנייה מאשרת.</div>
        <button class="big${A.armed === 'p' ? ' arm' : ''}" id="bulkP" ${np ? '' : 'disabled'}>${A.armed === 'p' ? 'לחץ שוב לאישור: ' : ''}אשר את כל הפיסוק שנותר בפסקה (${np})</button>
        <button class="big${A.armed === 'a' ? ' arm' : ''}" id="bulkA" ${here.length ? '' : 'disabled'}>${A.armed === 'a' ? 'לחץ שוב לאישור: ' : ''}אשר את כל מה שנותר בפסקה (${here.length})</button>
        <div id="bulkMsg" class="hint">${arg || ''}</div></div>
        <div class="box"><b>הערות טכניות</b> <span class="hint">(${tech.length} בסימן, לא נספרות ב"נותרו")</span>
        ${(techHere.length ? techHere : tech).map((it) => `<div class="note">• ${esc(it.note)}</div>`).join('') || '<div class="hint">אין</div>'}</div>
        <div class="box"><button class="big" id="toSum">סיכום וייצוא החלטות</button><button class="big" id="toList2">רשימת כל הפריטים</button><button class="big" id="toLoad">טען סימן אחר</button></div>
        <div class="hint">חלונית הסקירה · גרסה ${esc(window.PANE_VERSION || '')} · ${A.port instanceof WordPortFake.FakeDoc ? 'מצב הדגמה (מסמך דמה, שום דבר לא נכתב ב-Word)' : 'מחובר ל-Word'}</div>`;
    } else if (name === 'list') {
      const f = A.filter;
      const stOk = (s, it) => f.st === 'all' || (f.st === 'need' && NEED.has(s)) || (f.st === 'qq' && ((it.flags || []).includes('??') || it.kind === 'query') && NEED.has(s)) || s === f.st;
      const rows = A.order.filter((id) => stOk(st(id), item(id)) && (!f.kind || item(id).kind === f.kind));
      const kinds = [...new Set(A.order.map((id) => item(id).kind))];
      const pIx = new Map(pk.paragraphs.map((p, i) => [p.pid, i + 1]));
      h += `<h2>רשימה (${rows.length}) ${back}</h2>
        <div class="chips">${[['need', 'לטיפול'], ['qq', '?? ושאלות'], ['rejected', 'נדחו'], ['accepted', 'אושרו'], ['notfound', 'לא נמצאו'], ['changed', 'השתנו בקובץ'], ['all', 'הכול']].map(([k, l]) => `<button data-st="${k}" class="${f.st === k ? 'on' : ''}">${l}</button>`).join('')}</div>
        <div class="chips"><button data-kind="" class="${!f.kind ? 'on' : ''}">כל הסוגים</button>${kinds.map((k) => `<button data-kind="${k}" class="${f.kind === k ? 'on' : ''}">${KIND[k] || k}</button>`).join('')}</div>
        <div class="box" style="padding:2px 4px">${rows.map((id) => { const it = item(id), s = st(id); return `<div class="lrow" data-go="${id}"><span>פ' ${pIx.get(it.pid)}</span><span>${KIND[it.kind] || it.kind}</span><span class="t">${esc(listText(it))}</span><span>${STATE[s] || s}</span></div>`; }).join('') || '<div class="hint">אין פריטים בסינון הזה</div>'}</div>
        ${(pk.replies || []).length ? `<div class="box"><b>תשובות Claude להערות שלך (${pk.replies.length})</b>${pk.replies.map((r) => `<div class="note"><span class="hint">אתה:</span> ${esc(r.his_text)}<br><span class="hint">Claude:</span> ${esc(r.reply)}</div>`).join('')}</div>` : ''}`;
    } else if (name === 'sum') {
      const cnt = {}, byKind = {};
      A.order.forEach((id) => { const s = st(id), k = item(id).kind; cnt[s] = (cnt[s] || 0) + 1; byKind[k] = byKind[k] || {}; byKind[k][s] = (byKind[k][s] || 0) + 1; });
      const n = left();
      h += `<h2>סיכום ${back}</h2>
        ${n ? `<div class="box err">נשארו ${n} פריטים שלא הוחלט עליהם.</div>` : '<div class="box" style="color:var(--ok);font-weight:700">כל הפריטים טופלו.</div>'}
        <div class="box"><table class="cnt">${Object.entries(cnt).map(([s, c]) => `<tr><td>${STATE[s] || s}</td><td><b>${c}</b></td></tr>`).join('')}</table></div>
        <div class="box"><table class="cnt">${Object.entries(byKind).map(([k, o]) => `<tr><td>${KIND[k] || k}</td><td>${Object.entries(o).map(([s, c]) => `${STATE[s] || s} ${c}`).join(' · ')}</td></tr>`).join('')}</table></div>
        <div class="box"><button class="big" id="exp">ייצוא החלטות (קובץ להעלאה לשיחת הסימן)</button><div id="expMsg" class="hint"></div><textarea id="expT" readonly hidden></textarea></div>`;
    } else if (name === 'load') {
      h += `<h2>טעינת סימן ${A.S ? back : ''}</h2><div class="hint">לבחור את קובץ החבילה של הסימן ("סקירה - סימן … - סבב ….json") שנמסר עם הסימן.</div>
        <div class="drop" id="drop">לגרור לכאן את קובץ החבילה, או <label style="color:var(--acc);text-decoration:underline;cursor:pointer"><input type="file" id="pick" accept=".json" hidden>לבחור קובץ</label></div><div id="loadMsg" class="err"></div>`;
    }
    sc.innerHTML = h; sc.scrollTop = 0;
    const on = (id, fn) => { const el = $(id); if (el) el.onclick = fn; };
    on('scBack', closeScreen);
    on('toSum', () => screen('sum')); on('toList2', () => screen('list')); on('toLoad', () => screen('load'));
    const arm = (key, punctOnly) => async () => {
      if (A.armed !== key) { A.armed = key; return screen('more'); }
      A.armed = null; const r = await bulk(punctOnly); A.pid = null; screen('more', `נכתבו ${r.ok}${r.bad ? `, ${r.bad} לא נכתבו (ראו ברשימה)` : ''}.`);
    };
    on('repOk', () => {
      const t = $('repT').value.trim(); if (!t) return;
      const it = A.cur && item(A.cur);
      A.reports.push({ time: new Date().toISOString(), note: t, version: window.PANE_VERSION, siman: pk.siman, round: pk.round, item: A.cur || null, kind: it ? it.kind : null, op: it ? it.op : null,
        state: A.cur ? st(A.cur) : null, pid: A.pid, recent: A.hist.slice(-6), diag: A.port.diag ? A.port.diag.slice(-12) : [], images: (A.repImgs || []).slice() });
      A.repImgs = [];
      save(); screen('more', 'הדיווח נשמר (' + A.reports.length + ').');
    });
    const addImgs = async (files) => {
      A.repImgs = A.repImgs || [];
      for (const f of files) { if (f && /^image\//.test(f.type)) { try { A.repImgs.push(await shrinkImg(f)); } catch (e) { /* */ } } }
      const n = $('repImgN'); if (n) n.textContent = A.repImgs.length ? A.repImgs.length + ' תמונות מצורפות' : '';
    };
    { const el = $('repImg'); if (el) el.onchange = () => addImgs([...el.files]); }
    { const el = $('repT'); if (el) el.onpaste = (ev) => { const fs = [...((ev.clipboardData && ev.clipboardData.items) || [])].filter((x) => x.kind === 'file').map((x) => x.getAsFile()); if (fs.length) { ev.preventDefault(); addImgs(fs); } }; }
    on('repExp', async () => {
      const L = newReps(); if (!L.length) return;
      download(`דיווחים על התוסף - ${new Date().toISOString().replace('T', ' ').slice(0, 16).replace(':', '-')}.json`, JSON.stringify({ reports: L }, null, 1));
      L.forEach((r) => A.repDone.add(repKey(r)));
      try { await DB.put('reportsExported', [...A.repDone]); } catch (e) { /* */ }
      screen('more', 'יוצאו ' + L.length + ' דיווחים. בייצוא הבא ייצאו רק דיווחים חדשים.');
    });
    const setT = (id, k) => { const el = $(id); if (el) el.onchange = () => { A.settings[k] = el.checked; try { localStorage.setItem('review-pane-settings', JSON.stringify(A.settings)); } catch (e) { /* */ } if (A.port) A.port.highlightHeadings = A.settings.hl; if (k === 'mark' && !el.checked && A.port.mark) A.port.mark(null).catch(() => {}); }; };
    setT('setHl', 'hl'); setT('setMark', 'mark');
    on('bulkP', arm('p', true)); on('bulkA', arm('a', false));
    sc.querySelectorAll('[data-st]').forEach((b) => { b.onclick = () => { A.filter.st = b.dataset.st; screen('list'); }; });
    sc.querySelectorAll('[data-kind]').forEach((b) => { b.onclick = () => { A.filter.kind = b.dataset.kind || null; screen('list'); }; });
    sc.querySelectorAll('[data-go]').forEach((b) => { b.onclick = () => { sc.hidden = true; sc.innerHTML = ''; A.pid = null; go(b.dataset.go); }; });
    on('exp', () => {
      const data = exportDecisions(), txt = JSON.stringify(data, null, 1);
      const name = `החלטות - סימן ${pk.siman} - סבב ${pk.round}.json`;
      try {
        const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([txt], { type: 'application/json' })); a.download = name;
        document.body.appendChild(a); a.click(); a.remove();
        $('expMsg').textContent = `נשמר "${name}" בתיקיית ההורדות. אם לא נשמר, להעתיק את הטקסט שלמטה.`;
      } catch (e) { $('expMsg').textContent = 'ההורדה חסומה: להעתיק את הטקסט שלמטה.'; }
      $('expT').hidden = false; $('expT').value = txt;
    });
    if (name === 'load') {
      const take = (file) => { const fr = new FileReader(); fr.onload = () => { try { const pkg = JSON.parse(fr.result); sc.hidden = true; sc.innerHTML = ''; A.pid = null; A.times = {}; loadPkg(pkg).catch((e2) => { screen('load'); $('loadMsg').textContent = 'הטעינה נכשלה: ' + e2.message; }); } catch (e) { $('loadMsg').textContent = 'הקובץ אינו חבילת סקירה: ' + e.message; } }; fr.readAsText(file); };
      $('pick').onchange = (e) => e.target.files[0] && take(e.target.files[0]);
      const d = $('drop');
      d.ondragover = (e) => { e.preventDefault(); d.classList.add('over'); };
      d.ondragleave = () => d.classList.remove('over');
      d.ondrop = (e) => { e.preventDefault(); d.classList.remove('over'); e.dataTransfer.files[0] && take(e.dataTransfer.files[0]); };
    }
  }
  function download(name, txt) {
    try { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([txt], { type: 'application/json' })); a.download = name; document.body.appendChild(a); a.click(); a.remove(); return true; } catch (e) { return false; }
  }
  function exportDecisions() {
    const pk = A.P.pkg;
    return { format: 1, siman: pk.siman, round: pk.round, exported: new Date().toISOString(), pane_version: window.PANE_VERSION,
      items: pk.items.map((it) => { const s = st(it.id), lg = A.S.log[it.id] || {}; return { id: it.id, kind: it.kind, state: s === 'info' ? 'open' : s, time: A.times[it.id] || lg.time || null, why: lg.why || '', verdict: lg.verdict || null, text: s === 'other' ? (lg.text || null) : null }; }) };
  }
  function listText(it) {
    if (it.cluster) { const c = A.P.clusters.get(it.cluster); return c.orig + ' ← ' + C.outcomeText(c, A.P.items, 1 << c.items.indexOf(it.id)); }
    if (it.op === 'split_before') return '¶ ' + A.P.paras.get(it.pid).orig_tokens.slice(it.from, it.from + 3).join(' ');
    return it.op === 'none' ? it.note : it.new;
  }
  function empty(t) { $('text').innerHTML = `<div class="empty">${esc(t)}</div>`; }

  // ───────── הפעלה
  A.start = start; A.screen = screen; A.go = go; A.page = page;
  window.addEventListener('DOMContentLoaded', () => {});
  async function boot(inWord) {
    A.inWord = !!inWord;
    $('ver').textContent = 'גרסה ' + (window.PANE_VERSION || '');
    if (A.inWord) {
      try {
        const pk = await DB.get('cur');
        if (pk) { await loadPkg(pk); return; }
      } catch (e) { screen('load'); $('loadMsg').textContent = 'הסימן השמור לא נטען: ' + e.message; return; }
    }
    screen('load');
  }
  if (window.DEMO_PKG) start(window.DEMO_PKG);
  else if (window.Office && Office.onReady) Office.onReady((info) => boot(info && info.host === Office.HostType.Word));
  else boot(false);
}());
