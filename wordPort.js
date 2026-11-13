/* wordPort.js — שלב 0: בדיקות היסוד (probe) בלבד.
   כל הקריאות ל-Office.js מרוכזות כאן. רץ רק על עותק של הקובץ.
   הבדיקות כותבות פסקאות בדיקה בסוף המסמך, מתחת לשורת סימון, ולא נוגעות בטקסט הקיים
   (חוץ מקריאה בלבד של הפסקה שבה עומד הסמן). */
(function () {
  "use strict";

  const MARK = "— בדיקות חלונית הסקירה. אפשר למחוק מכאן ועד סוף המסמך —";
  const TAG = (n) => "(בדיקה " + n + ")";
  const TEXT_STYLE = "4_טקסט-רץ";
  const HEAD_STYLE = "7_כותרת-צד";
  const REF_STYLE = '8_מ"מ תו';

  const PARAS = {
    1: 'אמר רבא הלכה כרבי יהודה וכתב הרמב"ם בהל\' שבת פ״ב שמותר ' + TAG(1),
    2: "זה טקסט עם מחיקה במעקב באמצע המשפט " + TAG(2),
    3: "פסקה שלפניה תיכנס כותרת צד " + TAG(3),
    4: "וכן מבואר בגמרא שם ודו\"ק " + TAG(4),
    5: "כמבואר בגמרא ", // + מ"מ מעוצב + המשך, נבנה בהכנה
    6: "חלק ראשון של הפסקה. חלק שני של הפסקה " + TAG(6),
    7: "ועוד יש לומר דהנה מבואר בדברי הפוסקים שכל דבר שאין בו ממש אינו נקנה אלא בדרך הקנאה גמורה ולכן צריך עיון בדבריו שם היטב " + TAG(7),
  };

  const out = [];
  const log = (s) => out.push(s);
  const ok = (s) => log("✓ " + s);
  const bad = (s) => log("✗ " + s);
  const info = (s) => log("  " + s);
  const ms = (t) => Math.round(t) + " מ\"ש";
  const short = (s, n) => { s = String(s == null ? "" : s).replace(/\r/g, "¶").replace(/\n/g, "↵"); return s.length > (n || 80) ? s.slice(0, n || 80) + "…" : s; };
  const fmtDate = (d) => { try { const x = new Date(d); return x.toLocaleTimeString("he-IL", { hour12: false }) + " " + x.toLocaleDateString("he-IL"); } catch (e) { return String(d); } };

  async function step(title, fn) {
    log("");
    log("■ " + title);
    try { await fn(); }
    catch (e) { bad("שגיאה: " + (e && (e.code ? e.code + " – " : "") + (e.message || e))); if (e && e.debugInfo) info("debugInfo: " + short(JSON.stringify(e.debugInfo), 300)); }
  }

  async function findPara(ctx, n) {
    const res = ctx.document.body.search(TAG(n), { matchCase: true });
    res.load("items");
    await ctx.sync();
    if (!res.items.length) throw new Error("פסקת בדיקה " + n + " לא נמצאה");
    return res.items[res.items.length - 1].paragraphs.getFirst();
  }

  async function tcList(ctx, target) {
    const tcs = target.getTrackedChanges();
    tcs.load("items/type,items/text,items/author,items/date");
    await ctx.sync();
    return tcs.items.map((t) => ({ type: t.type, text: t.text, author: t.author, date: t.date }));
  }
  const tcLine = (t) => t.type + " «" + short(t.text, 40) + "» מחבר: " + t.author + " · " + fmtDate(t.date);

  async function tokens(ctx, para, trim) {
    const tr = para.getTextRanges([" "], trim);
    tr.load("items/text");
    await ctx.sync();
    return tr;
  }

  function apiSupport() {
    const rows = [];
    const sets = [["WordApi", ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]], ["WordApiDesktop", ["1.1", "1.2", "1.3", "1.4"]], ["WordApiHiddenDocument", ["1.5"]]];
    sets.forEach(([name, vers]) => {
      const yes = vers.filter((v) => { try { return Office.context.requirements.isSetSupported(name, v); } catch (e) { return false; } });
      rows.push(name + ": " + (yes.length ? yes.join(", ") : "אין"));
    });
    return rows;
  }

  /* ---------- הכנה: פסקאות בדיקה בסוף המסמך, כשהמעקב כבוי ---------- */
  async function setup(ctx, state) {
    const body = ctx.document.body;
    const prev = body.search(MARK, { matchCase: true });
    prev.load("items");
    await ctx.sync();
    if (prev.items.length) throw new Error("כבר יש במסמך פסקאות בדיקה מהרצה קודמת. לחץ קודם \"מחק את פסקאות הבדיקה\".");

    ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await ctx.sync();

    const styles = ctx.document.getStyles();
    styles.load("items/nameLocal,items/type");
    await ctx.sync();
    const names = styles.items.map((s) => s.nameLocal);
    state.hasText = names.includes(TEXT_STYLE);
    state.hasHead = names.includes(HEAD_STYLE);
    state.hasRef = names.includes(REF_STYLE);
    info("סגנונות במסמך: " + names.length + ". " + TEXT_STYLE + ": " + (state.hasText ? "יש" : "אין") + " · " + HEAD_STYLE + ": " + (state.hasHead ? "יש" : "אין") + " · " + REF_STYLE + ": " + (state.hasRef ? "יש" : "אין"));

    body.insertParagraph(MARK, "End");
    for (const n of [1, 2, 3, 4, 5, 6, 7]) {
      const p = body.insertParagraph(PARAS[n], "End");
      if (state.hasText) p.style = TEXT_STYLE;
      if (n === 5) {
        const ref = p.getRange("End").insertText("(ברכות ב, א)", "Before");
        if (state.hasRef) ref.style = REF_STYLE;
        p.getRange("End").insertText(" ודו\"ק " + TAG(5), "Before");
      }
    }
    await ctx.sync();
    ok("נכתבו 7 פסקאות בדיקה בסוף המסמך (המעקב כבוי בזמן ההכנה, כדי שלא ייחשבו שינוי)");
  }

  /* ---------- probe ---------- */
  async function probe() {
    out.length = 0;
    const state = {};
    log("דוח בדיקות היסוד — חלונית הסקירה " + (window.PANE_VERSION || ""));
    log("תאריך: " + fmtDate(new Date()));
    try {
      const d = Office.context.diagnostics;
      log("Word: " + d.host + " " + d.version + " · " + d.platform);
    } catch (e) { log("Word: לא ידוע"); }
    try { log("קובץ: " + (Office.context.document.url || "לא שמור")); } catch (e) { /* */ }
    apiSupport().forEach((r) => log("ממשקים · " + r));
    log("גודל החלונית: " + window.innerWidth + "×" + window.innerHeight + " · יחס פיקסלים " + window.devicePixelRatio);

    await Word.run(async (ctx) => {
      // (1) מעקב שינויים
      await step("1. מעקב שינויים: קריאה והדלקה", async () => {
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        state.origMode = ctx.document.changeTrackingMode;
        ok("מצב המעקב לפני הבדיקה: " + state.origMode);
      });

      await step("הכנה", async () => { await setup(ctx, state); });

      await step("1ב. הדלקת המעקב", async () => {
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        (ctx.document.changeTrackingMode === "TrackAll" ? ok : bad)("אחרי הדלקה: " + ctx.document.changeTrackingMode);
      });

      // (2) תיבות, פסיק, החלפה
      await step("2. חלוקה לתיבות, הוספת פסיק, החלפת מילה", async () => {
        const p = await findPara(ctx, 1);
        const tT = await tokens(ctx, p, true);
        const tF = await tokens(ctx, p, false);
        ok("תיבות (בלי רווחים): " + tT.items.length + ": " + tT.items.map((r) => "[" + r.text + "]").join(" "));
        info("תיבות (עם רווחים): " + tF.items.map((r) => "[" + r.text + "]").join(" "));
        tT.items[2].insertText(",", "End"); // אחרי "הלכה"
        await ctx.sync();
        p.load("text");
        let tcs = await tcList(ctx, p);
        await ctx.sync();
        info("אחרי פסיק: " + short(p.text, 120));
        info("שינויים במעקב: " + tcs.length);
        tcs.forEach((t) => info("· " + tcLine(t)));
        if (tcs.length === 1 && tcs[0].text === ",") ok("הפסיק נרשם כהוספה של תו אחד");
        else bad("הפסיק לא נרשם כהוספה של תו אחד בדיוק");
        tT.items[1].insertText("רבה", "Replace"); // רבא → רבה
        await ctx.sync();
        p.load("text");
        tcs = await tcList(ctx, p);
        info("אחרי החלפת רבא ברבה: " + short(p.text, 120));
        tcs.forEach((t) => info("· " + tcLine(t)));
        state.author = tcs.length ? tcs[tcs.length - 1].author : "";
        const dates = tcs.map((t) => fmtDate(t.date));
        ok("מחבר השינויים: " + state.author + " · שעות: " + dates.join(" | "));
      });

      // (3) מחיקה במעקב
      await step("3. פסקה עם מחיקה במעקב: טקסט, תיבות, טקסט נקי", async () => {
        const p = await findPara(ctx, 2);
        const hit = p.search("מחיקה ", { matchCase: true });
        hit.load("items");
        await ctx.sync();
        hit.items[0].delete();
        await ctx.sync();
        p.load("text");
        const cur = p.getRange("Whole").getReviewedText("Current");
        const orig = p.getRange("Whole").getReviewedText("Original");
        await ctx.sync();
        info("text: " + short(p.text, 120));
        info("getReviewedText(Current): " + short(cur.value, 120));
        info("getReviewedText(Original): " + short(orig.value, 120));
        const tr = await tokens(ctx, p, true);
        const tcs = p.getTrackedChanges();
        tcs.load("items/type");
        await ctx.sync();
        const dels = tcs.items.filter((t) => t.type === "Deleted").map((t) => t.getRange());
        const rel = [];
        tr.items.forEach((tok) => dels.forEach((d) => rel.push([tok, tok.compareLocationWith(d)])));
        await ctx.sync();
        const marks = tr.items.map((tok) => {
          const r = rel.filter((x) => x[0] === tok).map((x) => x[1].value);
          const isDel = r.some((v) => v === "Inside" || v === "Equal");
          return "[" + tok.text + (isDel ? " ✂" : "") + "]";
        });
        info("תיבות (✂ = בתוך מחיקה): " + marks.join(" "));
        const found = marks.filter((m) => m.includes("✂")).length;
        (found ? ok : bad)("זיהוי תיבות מחוקות דרך compareLocationWith: " + found + " נמצאו");
        (cur.value.indexOf("מחיקה") < 0 ? ok : bad)("הטקסט הנקי (Current) בלי המילה המחוקה");
      });

      await step("3ב. קריאה בלבד: הפסקה שבה עומד הסמן (מהטקסט האמיתי)", async () => {
        const p = ctx.document.getSelection().paragraphs.getFirst();
        p.load("text,style");
        const cur = p.getRange("Whole").getReviewedText("Current");
        await ctx.sync();
        if (p.text.indexOf("(בדיקה ") >= 0 || p.text === MARK) { info("הסמן עומד בפסקת בדיקה. כדי לבדוק טקסט אמיתי: לשים את הסמן בפסקה של הספר, רצוי עם מחיקה במעקב ועם גרשיים, ולהריץ שוב."); return; }
        const tr = await tokens(ctx, p, true);
        const tcs = await tcList(ctx, p);
        info("סגנון: " + p.style + " · אורך: " + p.text.length + " תווים · תיבות: " + tr.items.length);
        info("15 התיבות הראשונות: " + tr.items.slice(0, 15).map((r) => "[" + r.text + "]").join(" "));
        info("שינויים במעקב בפסקה: " + tcs.length + (tcs.length ? " (" + Array.from(new Set(tcs.map((t) => t.type + "/" + t.author))).join(", ") + ")" : ""));
        info("טקסט נקי (60 תווים ראשונים): " + short(cur.value, 60));
        ok("קריאת הפסקה האמיתית הצליחה, בלי כתיבה");
      });

      // (4) כותרת צד
      await step("4. הכנסת פסקה לפני פסקה, עם הסגנון " + HEAD_STYLE, async () => {
        const p = await findPara(ctx, 3);
        const np = p.insertParagraph("כותרת צד לבדיקה", "Before");
        np.style = HEAD_STYLE;
        await ctx.sync();
        np.load("style,text");
        const tcs = await tcList(ctx, np);
        (np.style === HEAD_STYLE ? ok : bad)("הסגנון שנקרא חזרה: " + np.style);
        tcs.forEach((t) => info("· " + tcLine(t)));
      });

      // (5) מראה מקום עם סגנון תו
      await step("5. הכנסת מראה מקום והחלת " + REF_STYLE, async () => {
        const p = await findPara(ctx, 4);
        const hit = p.search("שם", { matchCase: true });
        hit.load("items");
        await ctx.sync();
        const sp = hit.items[0].insertText(" ", "End");
        const ref = sp.insertText("(שבת קנו, א)", "After");
        ref.style = REF_STYLE;
        await ctx.sync();
        ref.load("style,text");
        p.load("text");
        const tcs = await tcList(ctx, p);
        (ref.style === REF_STYLE ? ok : bad)("סגנון הטווח שהוכנס: " + ref.style);
        info("הפסקה: " + short(p.text, 100));
        tcs.forEach((t) => info("· " + tcLine(t)));
        const types = Array.from(new Set(tcs.map((t) => t.type)));
        info("סוגי שינוי: " + types.join(", ") + (types.includes("Formatted") ? " ← נרשם גם כשינוי עיצוב" : " ← נרשם רק כהוספה"));
      });

      // (6) ירושת סגנון
      await step("6. הוספת תו בתוך ריצה שכבר בסגנון מ\"מ", async () => {
        const p = await findPara(ctx, 5);
        const hit = p.search("ברכות", { matchCase: true });
        hit.load("items");
        await ctx.sync();
        const r = hit.items[0];
        r.load("style");
        const add = r.insertText(" דף", "End");
        await ctx.sync();
        add.load("style,text");
        await ctx.sync();
        info("סגנון המילה \"ברכות\": " + r.style);
        (add.style === r.style ? ok : bad)("התוספת \"" + add.text + "\" קיבלה סגנון: " + add.style);
      });

      // (7) פיצול פסקה
      await step("7. פיצול פסקה", async () => {
        const p = await findPara(ctx, 6);
        const hit = p.search("חלק שני", { matchCase: true });
        hit.load("items");
        await ctx.sync();
        let how = "";
        try {
          hit.items[0].insertBreak(Word.BreakType.paragraph || "Paragraph", "Before");
          await ctx.sync();
          how = "insertBreak(paragraph)";
        } catch (e) {
          info("insertBreak(paragraph) נכשל: " + (e.message || e) + ". מנסה insertParagraph.");
          const p2 = await findPara(ctx, 6);
          const h2 = p2.search("חלק שני", { matchCase: true });
          h2.load("items");
          await ctx.sync();
          // ניסיון שני: תו מעבר פסקה
          h2.items[0].insertText("\n", "Before");
          await ctx.sync();
          how = "insertText(\\n)";
        }
        const a = ctx.document.body.search("חלק ראשון של הפסקה", { matchCase: true });
        const b = ctx.document.body.search("חלק שני של הפסקה", { matchCase: true });
        a.load("items"); b.load("items");
        await ctx.sync();
        const pa = a.items[a.items.length - 1].paragraphs.getFirst();
        const pb = b.items[b.items.length - 1].paragraphs.getFirst();
        pa.load("text,style"); pb.load("text,style");
        await ctx.sync();
        const split = pa.text !== pb.text;
        (split ? ok : bad)("פיצול דרך " + how + ": " + (split ? "שתי פסקאות" : "לא התפצל"));
        info("ראשונה: [" + pa.style + "] " + short(pa.text, 60));
        info("שנייה: [" + pb.style + "] " + short(pb.text, 60));
        const tcs = await tcList(ctx, pa);
        tcs.forEach((t) => info("· " + tcLine(t)));
      });

      // (8) זמנים
      await step("8. זמנים", async () => {
        let t = performance.now();
        const ps = ctx.document.body.paragraphs;
        ps.load("items/text,items/style");
        await ctx.sync();
        ok("קריאת כל פסקאות המסמך (" + ps.items.length + ", טקסט וסגנון): " + ms(performance.now() - t));
        t = performance.now();
        const all = ctx.document.body.getTrackedChanges();
        all.load("items/type");
        await ctx.sync();
        ok("קריאת כל השינויים במעקב במסמך (" + all.items.length + "): " + ms(performance.now() - t));
        let sum = 0;
        for (let i = 0; i < 10; i++) {
          t = performance.now();
          const sp = ctx.document.getSelection().paragraphs.getFirst();
          sp.load("text");
          await ctx.sync();
          sum += performance.now() - t;
        }
        ok("קריאת פסקת הסמן (ממוצע של 10): " + ms(sum / 10));
        const p = await findPara(ctx, 7);
        t = performance.now();
        const tr = await tokens(ctx, p, true);
        ok("חלוקה לתיבות של פסקה ארוכה (" + tr.items.length + " תיבות): " + ms(performance.now() - t));
        const times = [];
        for (let i = 0; i < 20 && i < tr.items.length - 2; i++) {
          t = performance.now();
          tr.items[i].insertText(",", "End");
          await ctx.sync();
          times.push(performance.now() - t);
        }
        const avg = times.reduce((a, b) => a + b, 0) / times.length;
        ok("כתיבת פסיק במעקב, " + times.length + " ברצף: ממוצע " + ms(avg) + ", הכי איטי " + ms(Math.max.apply(null, times)));
        const tcs = await tcList(ctx, p);
        info("שינויים בפסקה אחרי 20 פסיקים: " + tcs.length + " (מצופה 20; פחות = Word איחד הוספות סמוכות)");
      });

      // החזרת מצב המעקב
      await step("סיום: החזרת מצב המעקב כפי שהיה", async () => {
        if (state.origMode && state.origMode !== "TrackAll") {
          ctx.document.changeTrackingMode = state.origMode;
          await ctx.sync();
          info("המעקב הוחזר ל-" + state.origMode);
        } else info("המעקב נשאר דלוק, כמו שהיה.");
      });
    });
    return out.join("\n");
  }

  /* ---------- בדיקה 10: בחירה וגלילה ---------- */
  async function selectTest(which) {
    const t0 = performance.now();
    let before = 0, after = 0, label = "";
    await Word.run(async (ctx) => {
      const p1 = await findPara(ctx, 1);
      before = (await tcList(ctx, p1)).length;
      let target;
      if (which === "jump") { target = (await tokens(ctx, p1, true)).items[0]; label = "קפיצה לפסקאות הבדיקה"; }
      else if (which === "near") { const p = await findPara(ctx, 3); target = (await tokens(ctx, p, true)).items[1]; label = "מילה שכבר על המסך"; }
      else if (which === "near2") { const p = await findPara(ctx, 1); target = (await tokens(ctx, p, true)).items.slice(4, 7); label = "שלוש מילים על המסך"; }
      else { const p = ctx.document.body.paragraphs.getFirst(); target = (await tokens(ctx, p, true)).items[0]; label = "מילה רחוקה (תחילת המסמך)"; }
      if (Array.isArray(target)) target = target[0].expandTo(target[target.length - 1]);
      target.select("Select");
      await ctx.sync();
      after = (await tcList(ctx, p1)).length;
    });
    return { label, time: performance.now() - t0, changed: after !== before };
  }

  /* ---------- ניקוי ---------- */
  async function cleanup() {
    await Word.run(async (ctx) => {
      ctx.document.load("changeTrackingMode");
      const res = ctx.document.body.search(MARK, { matchCase: true });
      res.load("items");
      await ctx.sync();
      if (!res.items.length) throw new Error("לא נמצאו פסקאות בדיקה.");
      const mode = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      const start = res.items[0].paragraphs.getFirst().getRange("Start");
      const rng = start.expandTo(ctx.document.body.getRange("End"));
      rng.delete();
      await ctx.sync();
      // פסקה ריקה שאולי נשארה בסוף
      ctx.document.changeTrackingMode = mode;
      await ctx.sync();
    });
  }

  window.wordPort = { probe, selectTest, cleanup, MARK };
})();
