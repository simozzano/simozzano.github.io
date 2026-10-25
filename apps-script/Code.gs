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
const TOKEN = 'CAMBIA-QUESTO-TOKEN';
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
      const rows = Array.isArray(data) ? data : data.rows;
      if (!rows || !rows.length) return; // mai svuotare un foglio
      const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
      const values = rows.map(function (r) {
        const row = [];
        for (let i = 0; i < width; i++) {
          const v = i < r.length ? r[i] : null;
          if (v === null || v === undefined) row.push('');
          else if (typeof v === 'object' && v.date) { row.push(Utilities.parseDate(v.date, tz, 'yyyy-MM-dd')); }
          else row.push(v);
        }
        return row;
      });
      // 1) prima si scrive, e solo dopo si cancella l'eventuale coda: se la scrittura fallisce i dati vecchi restano
      if (sh.getMaxRows() < values.length) sh.insertRowsAfter(sh.getMaxRows(), values.length - sh.getMaxRows());
      if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
      const range = sh.getRange(1, 1, values.length, width);
      // 2) le convalide (menu a tendina) rifiuterebbero la scrittura: si tolgono un attimo e si rimettono,
      //    estendendo alle righe nuove la regola della riga 2
      const rules = range.getDataValidations();
      const rulesNew = rules.map(function (r, ri) { return r.map(function (v, ci) { return v || (ri >= 1 ? rules[1][ci] : null) || null; }); });
      range.clearDataValidations();
      (DATE_COLS[name] || []).forEach(function (c) { if (c <= width) sh.getRange(1, c, values.length, 1).setNumberFormat(DATE_FMT); });
      range.setValues(values);
      try { range.setDataValidations(rulesNew); } catch (err) { /* convalide non ripristinabili: i dati sono comunque salvati */ }
      // 3) coda: righe e colonne oltre i nuovi dati
      const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
      if (lastRow > values.length) sh.getRange(values.length + 1, 1, lastRow - values.length, lastCol).clearContent();
      if (lastCol > width) sh.getRange(1, width + 1, Math.min(lastRow, values.length), lastCol - width).clearContent();
    });
    SpreadsheetApp.flush();
    return reply_({ rev: revOf_(readAll_()) }, id);
  } catch (err) {
    return reply_({ error: 'Errore nello script: ' + (err && err.message ? err.message : err) + (err && err.stack ? ' | ' + String(err.stack).split('\n')[0] : '') }, id);
  } finally {
    lock.releaseLock();
  }
}
