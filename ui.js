/* ui.js — חיווט החלונית של שלב 0 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const S = { probe: "", sel: [], answers: {}, key: "", file: "", idb: "", dl: "", csp: [], size: "", inOffice: false };

  function setStatus(id, text, cls) { const el = $(id); el.textContent = text; el.className = "status" + (cls ? " " + cls : ""); }

  function report() {
    const L = [];
    L.push(S.probe || "(הבדיקות האוטומטיות עוד לא רצו)");
    L.push("");
    L.push("■ 10. בחירה וגלילה");
    S.sel.forEach((r) => L.push("  " + r.label + ": " + Math.round(r.time) + " מ\"ש · " + (r.changed ? "✗ נרשם שינוי בקובץ!" : "לא נרשם שינוי") + (r.err ? " · שגיאה: " + r.err : "")));
    document.querySelectorAll(".q").forEach((q) => L.push("  " + q.firstChild.textContent.trim() + " → " + (S.answers[q.dataset.q] || "לא נענה")));
    L.push("  מקש בחלונית אחרי בחירה: " + (S.key || "לא נלחץ"));
    L.push("");
    L.push("■ 9. קבצים");
    L.push("  בחירה/גרירה: " + (S.file || "לא נבדק"));
    L.push("  IndexedDB: " + (S.idb || "לא נבדק"));
    L.push("  הורדה: " + (S.dl || "לא נבדק"));
    L.push("");
    L.push("■ 11. גודל החלונית: " + S.size);
    L.push("■ CSP: " + (S.csp.length ? S.csp.join(" | ") : "אין חסימות"));
    $("report").value = L.join("\n");
  }

  function size() { S.size = window.innerWidth + "×" + window.innerHeight + " (יחס " + window.devicePixelRatio + ")"; $("size").textContent = "חלונית " + S.size; report(); }

  /* ---- שאלות כן/לא ---- */
  function buildQuestions() {
    document.querySelectorAll(".q").forEach((q) => {
      const box = document.createElement("div");
      box.className = "ans";
      ["כן", "לא", "לא בטוח"].forEach((v) => {
        const l = document.createElement("label");
        const i = document.createElement("input");
        i.type = "radio"; i.name = q.dataset.q; i.value = v;
        i.addEventListener("change", () => { S.answers[q.dataset.q] = v; report(); });
        l.appendChild(i); l.appendChild(document.createTextNode(v));
        box.appendChild(l);
      });
      q.appendChild(box);
    });
  }

  /* ---- IndexedDB ---- */
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open("review-pane-stage0", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("files");
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbPut(name, data) {
    const db = await idb();
    await new Promise((res, rej) => { const tx = db.transaction("files", "readwrite"); tx.objectStore("files").put({ name, data, at: new Date().toISOString() }, "last"); tx.oncomplete = res; tx.onerror = () => rej(tx.error); });
  }
  async function idbGet() {
    const db = await idb();
    return new Promise((res, rej) => { const tx = db.transaction("files", "readonly"); const g = tx.objectStore("files").get("last"); g.onsuccess = () => res(g.result); g.onerror = () => rej(g.error); });
  }
  async function takeFile(f, how) {
    try {
      const data = await f.text();
      S.file = "✓ " + f.name + " (" + f.size + " בתים) נקרא" + (how ? " (" + how + ")" : "");
      setStatus("fileStatus", S.file, "ok");
      await idbPut(f.name, data.slice(0, 200000));
      S.idb = "✓ נשמר ב-IndexedDB; לבדוק אחרי סגירת Word ופתיחה מחדש";
      setStatus("idbStatus", S.idb, "ok");
    } catch (e) { S.file = "✗ " + e.message; setStatus("fileStatus", S.file, "bad"); }
    report();
  }

  function download(name, text) {
    try {
      const blob = new Blob(["﻿" + text], { type: "text/plain;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click(); a.remove();
      return true;
    } catch (e) { return e.message; }
  }

  async function copy() {
    const t = $("report");
    try { await navigator.clipboard.writeText(t.value); return true; }
    catch (e) { t.focus(); t.select(); try { return document.execCommand("copy"); } catch (e2) { return false; } }
  }

  function wire() {
    buildQuestions();
    $("ver").textContent = "גרסה " + (window.PANE_VERSION || "?");
    window.addEventListener("resize", size); size();
    document.addEventListener("securitypolicyviolation", (e) => { S.csp.push(e.violatedDirective + " " + e.blockedURI); report(); });
    document.addEventListener("keydown", (e) => { if (e.target.tagName === "TEXTAREA") return; S.key = "✓ התקבל (" + (e.key === " " ? "רווח" : e.key) + ")"; $("keyResult").textContent = S.key; report(); });

    $("isCopy").addEventListener("change", () => { $("runProbe").disabled = !($("isCopy").checked && S.inOffice); $("safety").classList.toggle("okd", $("isCopy").checked); });

    $("runProbe").addEventListener("click", async () => {
      $("runProbe").disabled = true;
      setStatus("probeStatus", "רץ… (עד דקה בקובץ גדול)");
      try {
        S.probe = await window.wordPort.probe();
        const bads = (S.probe.match(/^✗/gm) || []).length;
        setStatus("probeStatus", bads ? "הסתיים, עם " + bads + " בעיות. הפרטים בדוח למטה." : "הסתיים. כל הבדיקות האוטומטיות עברו.", bads ? "bad" : "ok");
      } catch (e) { S.probe = "✗ הבדיקות נעצרו: " + (e.message || e); setStatus("probeStatus", S.probe, "bad"); }
      $("runProbe").disabled = false;
      report();
    });

    document.querySelectorAll("[data-sel]").forEach((b) => b.addEventListener("click", async () => {
      try {
        const r = await window.wordPort.selectTest(b.dataset.sel);
        S.sel.push(r);
        setStatus("selStatus", r.label + ": " + Math.round(r.time) + " מ\"ש" + (r.changed ? " · נרשם שינוי!" : ""), r.changed ? "bad" : "ok");
      } catch (e) { S.sel.push({ label: b.textContent, time: 0, err: e.message || String(e) }); setStatus("selStatus", "שגיאה: " + (e.message || e), "bad"); }
      report();
    }));

    const drop = $("drop");
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); const f = e.dataTransfer.files[0]; if (f) takeFile(f, "בגרירה"); });
    $("filePick").addEventListener("change", (e) => { const f = e.target.files[0]; if (f) takeFile(f, "בבחירה"); });

    $("dl").addEventListener("click", () => { const r = download("בדיקת הורדה - חלונית הסקירה.txt", "בדיקת הורדה מחלונית הסקירה, " + new Date().toLocaleString("he-IL")); S.dl = r === true ? "הלחיצה עברה בלי שגיאה (לבדוק בתיקיית ההורדות)" : "✗ " + r; report(); });
    $("copyReport").addEventListener("click", async () => { const r = await copy(); setStatus("copyStatus", r ? "הדוח הועתק. להדביק אותו בשיחה." : "ההעתקה נחסמה: לסמן את הטקסט בתיבה ולהעתיק ב-Ctrl+C.", r ? "ok" : "bad"); });
    $("saveReport").addEventListener("click", () => download("דוח בדיקות היסוד - חלונית הסקירה.txt", $("report").value));
    $("cleanup").addEventListener("click", async () => {
      try { await window.wordPort.cleanup(); setStatus("cleanStatus", "פסקאות הבדיקה נמחקו.", "ok"); }
      catch (e) { setStatus("cleanStatus", "שגיאה: " + (e.message || e), "bad"); }
    });

    idbGet().then((r) => { S.idb = r ? "✓ נמצא קובץ שנשמר: " + r.name + " (" + new Date(r.at).toLocaleString("he-IL") + ")" : "אין עדיין קובץ שמור"; setStatus("idbStatus", S.idb, r ? "ok" : ""); report(); })
      .catch((e) => { S.idb = "✗ " + e.message; setStatus("idbStatus", S.idb, "bad"); report(); });
    report();
  }

  function officeReady(info) {
    S.inOffice = !!(info && info.host);
    try { $("docname").textContent = "קובץ: " + (Office.context.document.url || "לא שמור"); } catch (e) { /* */ }
    $("runProbe").disabled = !($("isCopy").checked && S.inOffice);
  }

  document.addEventListener("DOMContentLoaded", () => {
    wire();
    if (window.Office && Office.onReady) Office.onReady(officeReady);
    else $("docname").textContent = "מצב תצוגה בדפדפן: אין חיבור ל-Word";
  });
})();
