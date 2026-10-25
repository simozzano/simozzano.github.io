/**
 * Backend di "Conti di casa" per Google Sheets.
 *
 * Installazione (una volta sola, dal Google Sheet con i fogli Spese, Storico,
 * Categorie, Impostazioni, Persone, PesiAnnuali, Riassunto):
 *   1. Estensioni → Apps Script, incolla questo file al posto del contenuto.
 *   2. Cambia TOKEN con una stringa lunga e segreta (la stessa andrà nell'app).
 *   3. Distribuisci → Nuova distribuzione → tipo "App web":
 *        Esegui come: Me   |   Chi ha accesso: Chiunque
 *   4. Copia l'URL che termina con /exec e incollalo nell'app con il token.
 * Dopo ogni modifica a questo file: Distribuisci → Gestisci distribuzioni →
 * modifica → Nuova versione.
 */
const TOKEN = 'CAMBIA-QUESTA-STRINGA';
const SHEETS = ['Spese', 'Storico', 'Categorie', 'Impostazioni', 'Persone', 'PesiAnnuali'];
const DATE_COLS = { Spese: [1], Storico: [2], Persone: [4] }; // colonne (1-based) con date
const DATE_FMT = 'dd/mm/yyyy';

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheets = {};
  SHEETS.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const rows = sh.getDataRange().getValues().map(function (r) {
      return r.map(function (v) {
        if (v instanceof Date) return { date: Utilities.formatDate(v, tz, 'yyyy-MM-dd') };
        return v === '' ? null : v;
      });
    });
    sheets[name] = { state: sh.isSheetHidden() ? 'hidden' : 'visible', rows: rows };
  });
  return sheets;
}

function revOf_(sheets) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(sheets));
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) return out_({ error: 'Token non valido' });
  const sheets = readAll_();
  return out_({ rev: revOf_(sheets), sheets: sheets });
}

// Risposta alla POST: pagina HTML che rimanda il risultato alla pagina dell'app con postMessage
// (l'app invia un form in un iframe nascosto, così non dipende da CORS).
function reply_(obj, id) {
  obj.gsId = id;
  const json = JSON.stringify(obj).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<script>window.top.postMessage(' + json + ',"*");</script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.parameter.payload || e.postData.contents); } catch (err) { return reply_({ error: 'Richiesta non valida' }, null); }
  const id = body.id;
  if (body.token !== TOKEN) return reply_({ error: 'Token non valido' }, id);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return reply_({ error: 'Foglio occupato, riprova' }, id);
  try {
    const cur = revOf_(readAll_());
    if (body.baseRev && body.baseRev !== cur) return reply_({ conflict: true, rev: cur }, id);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    SHEETS.forEach(function (name) {
      const data = body.sheets && body.sheets[name];
      const sh = ss.getSheetByName(name);
      if (!data || !sh) return;
      const width = data.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
      const values = data.rows.map(function (r) {
        const row = [];
        for (let i = 0; i < width; i++) {
          const v = i < r.length ? r[i] : null;
          if (v === null || v === undefined) row.push('');
          else if (typeof v === 'object' && v.date) { row.push(Utilities.parseDate(v.date, tz, 'yyyy-MM-dd')); }
          else row.push(v);
        }
        return row;
      });
      sh.getDataRange().clearContent();
      if (values.length) {
        const range = sh.getRange(1, 1, values.length, width);
        (DATE_COLS[name] || []).forEach(function (c) { if (c <= width) sh.getRange(1, c, values.length, 1).setNumberFormat(DATE_FMT); });
        range.setValues(values);
      }
    });
    SpreadsheetApp.flush();
    return reply_({ rev: revOf_(readAll_()) }, id);
  } finally {
    lock.releaseLock();
  }
}
