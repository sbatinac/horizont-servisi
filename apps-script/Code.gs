/**
 * Horizont LHB d.o.o. – servisi
 * Prima prijave s web stranice, zapisuje ih u Google tablicu i šalje obavijest mailom.
 *
 * Postavljanje: vidi README.md → "Obrazac za prijavu".
 */

// Adresa na koju stižu obavijesti o novim prijavama.
const NOTIFY_EMAIL = 'horizont.servisi@gmail.com';
// Naziv lista u tablici.
const SHEET_NAME = 'Prijave';
// List za kratke upite s kontakt forme.
const SHEET_UPITI = 'Upiti';
// List s brojem posjeta po danu.
const SHEET_POSJETE = 'Posjete';
// Lozinka za admin stranicu. Upiši svoju ovdje u Apps Scriptu, NE u GitHub repo.
const ADMIN_KEY = 'UPISI_LOZINKU';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const upit = data._vrsta === 'upit';
    delete data._vrsta;
    const sheet = getSheet_(upit ? SHEET_UPITI : SHEET_NAME);
    const row = appendRow_(sheet, data);
    notify_(data, row, upit);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// GET: ?a=hit broji posjet, ?a=admin&k=LOZINKA vraća sve podatke za admin stranicu.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a === 'hit') return json_(hit_());
  if (p.a === 'admin') {
    if (!ADMIN_KEY || ADMIN_KEY === 'UPISI_LOZINKU' || p.k !== ADMIN_KEY) {
      Utilities.sleep(800);
      return json_({ ok: false, error: 'Pogrešna lozinka.' });
    }
    return json_({ ok: true, prijave: table_(SHEET_NAME), upiti: table_(SHEET_UPITI), posjete: table_(SHEET_POSJETE), tablica: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
  }
  return json_({ ok: true, servis: 'Horizont LHB – prijave' });
}

// Broj posjeta po danu (jedan posjet = jedna sesija posjetitelja).
function hit_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(SHEET_POSJETE);
    if (!sh.getLastRow()) { sh.appendRow(['Datum', 'Posjete']); sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#FEE503'); sh.setFrozenRows(1); }
    const today = Utilities.formatDate(new Date(), 'Europe/Zagreb', 'yyyy-MM-dd');
    const last = sh.getLastRow();
    if (last > 1 && sh.getRange(last, 1).getDisplayValue() === today) {
      const c = sh.getRange(last, 2); c.setValue(Number(c.getValue() || 0) + 1);
    } else {
      sh.appendRow(["'" + today, 1]);
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// Cijeli list kao { headers, rows } (prikazane vrijednosti, najnoviji redovi zadnji).
function table_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || !sh.getLastRow()) return { headers: [], rows: [] };
  const v = sh.getDataRange().getDisplayValues();
  return { headers: v[0], rows: v.slice(1) };
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
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
  // Brojevi telefona i slično ostaju tekst (da se ne izgubi početna nula).
  const row = headers.map(h => {
    const v = byLabel[h] !== undefined ? byLabel[h] : '';
    return (typeof v === 'string' && /^[+0][\d\s\/().-]*$/.test(v)) ? "'" + v : v;
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

// "Ime i prezime [ime]" → "Ime i prezime"
function label_(k) {
  return String(k).replace(/\s*\[[^\]]+\]$/, '');
}

function notify_(data, row, upit) {
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL.indexOf('@') === -1) return;
  const g = k => {
    const key = Object.keys(data).find(x => x.indexOf('[' + k + ']') !== -1);
    return key ? data[key] : '';
  };
  const uredaj = g('uredaj') + (g('tip_racunala') ? ' · ' + g('tip_racunala') : '');
  const subject = upit
    ? 'Novi upit s web stranice: ' + (g('ime') || 'bez imena')
    : 'Nova prijava servisa: ' + (g('ime') || 'bez imena') + ' – ' + (uredaj || 'uređaj');
  const lines = Object.keys(data)
    .filter(k => data[k] !== '' && data[k] !== undefined)
    .map(k => '<tr><td style="padding:4px 12px 4px 0;color:#666;vertical-align:top">' + esc_(label_(k)) +
              '</td><td style="padding:4px 0">' + esc_(data[k]) + '</td></tr>').join('');
  const html = '<h2 style="font-family:sans-serif">' + (upit ? 'Novi upit' : 'Nova prijava servisa') + '</h2>' +
    '<p style="font-family:sans-serif">Redak ' + row + ' u listu <a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">' + (upit ? SHEET_UPITI : SHEET_NAME) + '</a>.</p>' +
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
