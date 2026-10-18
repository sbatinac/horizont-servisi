/**
 * Horizont LHB d.o.o. – servisi
 * Prima prijave s web stranice, zapisuje ih u Google tablicu i šalje obavijest mailom.
 *
 * Postavljanje: vidi README.md → "Obrazac za prijavu".
 */

// Adresa na koju stižu obavijesti o novim prijavama.
const NOTIFY_EMAIL = 'UPISI_SVOJ_EMAIL';
// Naziv lista u tablici.
const SHEET_NAME = 'Prijave';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const row = appendRow_(sheet, data);
    notify_(data, row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Brza provjera da je skripta objavljena: otvori /exec link u pregledniku.
function doGet() {
  return json_({ ok: true, servis: 'Horizont LHB – prijave' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

// Stupci se dodaju sami: novo pitanje na stranici = novi stupac u tablici.
function appendRow_(sheet, data) {
  const keys = Object.keys(data);
  let headers = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  const missing = keys.filter(k => headers.indexOf(label_(k)) === -1);
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing.map(label_)]);
    headers = headers.concat(missing.map(label_));
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FEE503');
    sheet.setFrozenRows(1);
  }
  const byLabel = {};
  keys.forEach(k => byLabel[label_(k)] = data[k]);
  const row = headers.map(h => byLabel[h] !== undefined ? byLabel[h] : '');
  sheet.appendRow(row);
  return sheet.getLastRow();
}

// "Ime i prezime [ime]" → "Ime i prezime"
function label_(k) {
  return String(k).replace(/\s*\[[^\]]+\]$/, '');
}

function notify_(data, row) {
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL.indexOf('@') === -1) return;
  const g = k => {
    const key = Object.keys(data).find(x => x.indexOf('[' + k + ']') !== -1);
    return key ? data[key] : '';
  };
  const uredaj = g('uredaj') + (g('tip_racunala') ? ' · ' + g('tip_racunala') : '');
  const subject = 'Nova prijava servisa: ' + (g('ime') || 'bez imena') + ' – ' + (uredaj || 'uređaj');
  const lines = Object.keys(data)
    .filter(k => data[k] !== '' && data[k] !== undefined)
    .map(k => '<tr><td style="padding:4px 12px 4px 0;color:#666;vertical-align:top">' + esc_(label_(k)) +
              '</td><td style="padding:4px 0">' + esc_(data[k]) + '</td></tr>').join('');
  const html = '<h2 style="font-family:sans-serif">Nova prijava servisa</h2>' +
    '<p style="font-family:sans-serif">Redak ' + row + ' u tablici <a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">Prijave</a>.</p>' +
    '<table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">' + lines + '</table>';
  const replyTo = g('email');
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, htmlBody: html, replyTo: replyTo || undefined });
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
