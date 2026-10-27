/**
 * Daily Huddle – Wertung der Styrian Studs
 * Dieses Skript speichert die Ergebnisse in der Tabelle "Ergebnisse"
 * und liefert die Bestenliste an das Spiel.
 *
 * Einrichtung: einmal die Funktion "setup" ausführen,
 * danach als Web-App bereitstellen (Ausführen als: ich, Zugriff: jeder).
 */

const SHEET = 'Ergebnisse';
const SALT = 'studs2005';

// Ergebnisse aus dem Aufwärmen auf claude.ai (werden bei setup übernommen)
const IMPORT = [
  { d: '2026-10-25', num: '86', n: "Semmel", p: 331, s: 63 },
  { d: '2026-10-25', num: '41', n: "Chris", p: 293, s: 62 },
  { d: '2026-10-25', num: '95', n: "Tobi Doc Ock", p: 322, s: 65 },
  { d: '2026-10-26', num: '86', n: "Semmel", p: 353, s: 44 },
  { d: '2026-10-26', num: '43', n: "Obstla", p: 252, s: 113 },
  { d: '2026-10-26', num: '41', n: "Chris", p: 252, s: 80 },
  { d: '2026-10-26', num: '11', n: "Sunshine", p: 283, s: 61 },
  { d: '2026-10-26', num: '82', n: "Karlos", p: 285, s: 69 },
  { d: '2026-10-26', num: '49', n: "Clesch", p: 257, s: 75 },
  { d: '2026-10-26', num: '28', n: "Dave", p: 253, s: 204 },
  { d: '2026-10-26', num: '26', n: "Gasser", p: 300, s: 118 },
  { d: '2026-10-26', num: '33', n: "Justin", p: 396, s: 200 }
];

function setup() {
  const sh = sheet_();
  IMPORT.forEach(function (r) {
    if (!exists_(sh, r.d, r.num)) {
      sh.appendRow([new Date(), "'" + r.d, "'" + r.num, r.n, r.p, '', '', '', '', r.s]);
    }
  });
  Logger.log('Fertig! Die Tabelle "' + SHEET + '" ist bereit.');
}

function doGet(e) {
  const month = String((e && e.parameter && e.parameter.month) || '').slice(0, 7);
  const rows = sheet_().getDataRange().getValues().slice(1);
  const entries = rows
    .map(function (r) {
      return { d: day_(r[1]), num: String(r[2]), n: String(r[3]), p: Number(r[4]), s: Number(r[9]) };
    })
    .filter(function (x) { return !month || x.d.slice(0, 7) === month; });
  return json_({ ok: true, entries: entries });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    const num = String(parseInt(b.num, 10));
    const n = String(b.n || '').trim().slice(0, 18);
    const y = Number(b.y), m = Number(b.m), w = Number(b.w), h = Number(b.h), s = Number(b.s);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(b.d) || num === 'NaN' || +num < 0 || +num > 99 || !n) {
      return json_({ ok: false, reason: 'invalid', msg: 'Ungültige Angaben' });
    }
    if ([y, m, h].some(function (v) { return !(v >= 0 && v <= 100); }) || [0, 40, 70, 100].indexOf(w) < 0 || !(s > 0)) {
      return json_({ ok: false, reason: 'invalid', msg: 'Ungültige Punkte' });
    }
    if (b.c !== hash_(b.d, num, n, y, m, w, h, s)) {
      return json_({ ok: false, reason: 'code', msg: 'Prüfcode passt nicht' });
    }
    // Nur Ergebnisse von gestern, heute oder morgen (Zeitzonen-Puffer)
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const ok = [-1, 0, 1].map(function (k) {
      return Utilities.formatDate(new Date(Date.now() + k * 864e5), tz, 'yyyy-MM-dd');
    });
    if (ok.indexOf(b.d) < 0) {
      return json_({ ok: false, reason: 'date', msg: 'Falsches Datum' });
    }
    const sh = sheet_();
    if (exists_(sh, b.d, num)) {
      return json_({ ok: false, reason: 'dup' });
    }
    sh.appendRow([new Date(), "'" + b.d, "'" + num, n, y + m + w + h, y, m, w, h, s]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, reason: 'error', msg: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Hilfsfunktionen ----------
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Eingetragen', 'Datum', 'Nummer', 'Name', 'Punkte', 'Wann', 'Wo', 'Wer', 'Wie viel', 'Sekunden']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function exists_(sh, d, num) {
  const rows = sh.getDataRange().getValues().slice(1);
  return rows.some(function (r) { return day_(r[1]) === d && String(r[2]) === String(num); });
}

function day_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function norm_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function hash_(d, num, n, y, m, w, h, s) {
  const str = [d, num, norm_(n), y, m, w, h, s, SALT].join('|');
  let x = 2166136261;
  for (const ch of str) {
    x ^= ch.codePointAt(0);
    x = Math.imul(x, 16777619) >>> 0;
  }
  return x.toString(36).slice(-5);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
