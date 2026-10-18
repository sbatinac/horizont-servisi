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
// Drive: mapa galerije na stranici (Slike) i naziv mape za arhivu naloga.
const GALERIJA_ID = '1WI5VmXP9w2IgRyymJDDapsC8MQkkZnyD';
const GALERIJA_PODMAPA = 'Prije i poslije';
const ARHIVA_NAZIV = 'Horizont – Arhiva naloga';
// Podaci za mailove klijentima.
const SERVIS_NAZIV = 'Horizont LHB d.o.o. – servisi';
const SERVIS_TEL = '091/358 4122';
const WEB = 'https://sbatinac.github.io/horizont-servisi/';
const ADMIN_URL = WEB + 'admin.html';
// Lozinka za admin stranicu. Upiši svoju ovdje u Apps Scriptu, NE u GitHub repo.
const ADMIN_KEY = 'UPISI_LOZINKU';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data._vrsta === 'admin') return json_(admin_(data));
    const upit = data._vrsta === 'upit';
    delete data._vrsta;
    const sheet = getSheet_(upit ? SHEET_UPITI : SHEET_NAME);
    const nextRow = Math.max(sheet.getLastRow(), 1) + 1;
    const nr = upit ? '' : 'RN-' + Utilities.formatDate(new Date(), TZ, 'yyyy') + '-' + ('000' + (nextRow - 1)).slice(-4);
    if (nr) data['Broj naloga [broj_naloga]'] = nr;
    const row = appendRow_(sheet, data);
    notify_(data, row, upit);
    try { potvrdaKlijentu_(data, nr, upit); } catch (e) { console.warn(e); }
    return json_({ ok: true, nr: nr });
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
  if (p.a === 'status') return json_(statusLookup_(p.nr, p.tel));
  if (p.a === 'foto') {
    if (!ADMIN_KEY || ADMIN_KEY === 'UPISI_LOZINKU' || p.k !== ADMIN_KEY) return json_({ ok: false, error: 'Pogrešna lozinka.' });
    try { const b = DriveApp.getFileById(p.id).getBlob(); return json_({ ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) }); }
    catch (err) { return json_({ ok: false, error: 'Slika nije dostupna.' }); }
  }
  if (p.a === 'admin') {
    if (!ADMIN_KEY || ADMIN_KEY === 'UPISI_LOZINKU' || p.k !== ADMIN_KEY) {
      Utilities.sleep(800);
      return json_({ ok: false, error: 'Pogrešna lozinka.' });
    }
    return json_({ ok: true, prijave: table_(SHEET_NAME), upiti: table_(SHEET_UPITI), posjete: table_(SHEET_POSJETE), tablica: SpreadsheetApp.getActiveSpreadsheet().getUrl(), jutro: jutroAktivno_() });
  }
  return json_({ ok: true, servis: 'Horizont LHB – prijave' });
}

// ===== ADMIN AKCIJE (status, satnica, termin u kalendaru) =====
const CIJENA_SATA = 55;
const TZ = 'Europe/Zagreb';

function admin_(d) {
  if (!ADMIN_KEY || ADMIN_KEY === 'UPISI_LOZINKU' || d.k !== ADMIN_KEY) { Utilities.sleep(800); return { ok: false, error: 'Pogrešna lozinka.' }; }
  if (d.akcija === 'jutro') { postaviJutarnjiMail_(!!d.on); return { ok: true, jutro: jutroAktivno_() }; }
  if (d.akcija === 'jutro_test') { jutarnjiSazetak(); return { ok: true }; }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(d.list === 'upiti' ? SHEET_UPITI : SHEET_NAME);
  const row = Number(d.red);
  if (!sh || !row || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'Redak ne postoji. Osvježi stranicu.' };
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const ci = H.indexOf('Ime i prezime');
  if (ci > -1 && d.ime !== undefined && String(sh.getRange(row, ci + 1).getDisplayValue()) !== String(d.ime)) {
    return { ok: false, error: 'Tablica se promijenila u međuvremenu. Osvježi stranicu i pokušaj ponovno.' };
  }
  if (d.akcija === 'update') {
    const f = Object.assign({}, d.fields || {});
    if (f['Sati rada'] !== undefined && f['Sati rada'] !== '') {
      const h = Number(String(f['Sati rada']).replace(',', '.'));
      if (isNaN(h) || h < 0) return { ok: false, error: 'Neispravan broj sati.' };
      f['Sati rada'] = h; f['Iznos (€)'] = Math.round(h * CIJENA_SATA * 100) / 100;
    }
    setFields_(sh, row, f);
    return { ok: true };
  }
  if (d.akcija === 'termin') {
    const start = Utilities.parseDate(d.start, TZ, "yyyy-MM-dd'T'HH:mm");
    const end = new Date(start.getTime() + (Number(d.trajanje) || 60) * 60000);
    const get = h => { const i = H.indexOf(h); return i > -1 ? sh.getRange(row, i + 1).getDisplayValue() : ''; };
    const title = 'Servis: ' + (get('Ime i prezime') || 'klijent') + (get('Vrsta uređaja') ? ' – ' + get('Vrsta uređaja') : '');
    const desc = [['Telefon', get('Kontakt broj mobitela') || get('Telefon')], ['E-pošta', get('E-pošta')], ['Uređaj', [get('Vrsta uređaja'), get('Moje računalo je'), get('Točan model uređaja')].filter(String).join(' · ')],
      ['Usluga', get('Koju uslugu trebate?')], ['Kvar', get('Opišite problem ili kvar svojim riječima') || get('Poruka')], ['Predaja', get('Kako želite predati uređaj?')],
      ['Napomena', d.napomena || ''], ['Redak u tablici', row]].filter(x => x[1]).map(x => x[0] + ': ' + x[1]).join('\n');
    const loc = [get('Adresa'), get('Poštanski broj'), get('Grad / mjesto')].filter(String).join(', ');
    const cal = CalendarApp.getDefaultCalendar();
    let ev = null; const oldId = get('Kalendar ID');
    if (oldId) { try { ev = cal.getEventById(oldId); } catch (e) {} }
    if (ev) { ev.setTime(start, end); ev.setTitle(title); ev.setDescription(desc); ev.setLocation(loc); }
    else ev = cal.createEvent(title, start, end, { description: desc, location: loc });
    ev.addPopupReminder(60);
    setFields_(sh, row, { 'Termin': "'" + Utilities.formatDate(start, TZ, 'yyyy-MM-dd HH:mm'), 'Trajanje (min)': Number(d.trajanje) || 60, 'Kalendar ID': ev.getId(), 'Status': d.status || 'Zakazano' });
    return { ok: true, id: ev.getId() };
  }
  if (d.akcija === 'otkazi') {
    const i = H.indexOf('Kalendar ID');
    const id = i > -1 ? sh.getRange(row, i + 1).getDisplayValue() : '';
    if (id) { try { const ev = CalendarApp.getDefaultCalendar().getEventById(id); if (ev) ev.deleteEvent(); } catch (e) {} }
    setFields_(sh, row, { 'Termin': '', 'Trajanje (min)': '', 'Kalendar ID': '', 'Status': 'Novo' });
    return { ok: true };
  }
  if (d.akcija === 'slika' || d.akcija === 'galerija' || d.akcija === 'brisi_sliku' || d.akcija === 'pdf') return media_(sh, row, H, d);
  return { ok: false, error: 'Nepoznata akcija.' };
}

// ===== SLIKE, GALERIJA I ARHIVA NALOGA =====
function cell_(sh, row, H, name) { const i = H.indexOf(name); return i > -1 ? sh.getRange(row, i + 1).getDisplayValue() : ''; }

function arhiva_() {
  const it = DriveApp.getFoldersByName(ARHIVA_NAZIV);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ARHIVA_NAZIV);
}

function folderFor_(sh, row, H, d) {
  const id = cell_(sh, row, H, 'Mapa ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const m = String(cell_(sh, row, H, 'vrijeme')).match(/(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})/);
  const datum = m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const name = [d.nr || ('RN-' + row), datum, cell_(sh, row, H, 'Ime i prezime')].filter(String).join(' – ');
  const f = arhiva_().createFolder(name);
  setFields_(sh, row, { 'Mapa ID': f.getId(), 'Mapa': f.getUrl() });
  return f;
}

function galerija_() {
  const g = DriveApp.getFolderById(GALERIJA_ID);
  const it = g.getFoldersByName(GALERIJA_PODMAPA);
  return it.hasNext() ? it.next() : g.createFolder(GALERIJA_PODMAPA);
}

function trash_(id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }

function media_(sh, row, H, d) {
  let S = {}; try { S = JSON.parse(cell_(sh, row, H, 'Slike') || '{}'); } catch (e) {}
  if (d.akcija === 'pdf') {
    const f = folderFor_(sh, row, H, d);
    trash_(cell_(sh, row, H, 'PDF ID'));
    const pdf = Utilities.newBlob(d.html, 'text/html', 'nalog.html').getAs('application/pdf').setName((d.nr || 'Radni nalog') + '.pdf');
    const file = f.createFile(pdf);
    setFields_(sh, row, { 'PDF ID': file.getId(), 'PDF': file.getUrl() });
    return { ok: true, url: file.getUrl() };
  }
  const sm = String(d.slot).match(/^(prije|poslije)(\d{1,3})$/);
  if (!sm) return { ok: false, error: 'Nepoznato mjesto za sliku.' };
  const cur = S[d.slot] || {};
  const label = (d.nr || 'RN') + ' – ' + (sm[1] === 'prije' ? 'prije servisa' : 'poslije servisa') + ' – slika ' + sm[2];
  if (d.akcija === 'slika') {
    const f = folderFor_(sh, row, H, d);
    trash_(cur.id); trash_(cur.gid);
    const blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.mime || 'image/jpeg', label + '.jpg');
    const file = f.createFile(blob);
    S[d.slot] = { id: file.getId() };
    if (d.gal) { const c = file.makeCopy(label + '.jpg', galerija_()); S[d.slot].gid = c.getId(); }
  } else if (d.akcija === 'galerija') {
    if (!cur.id) return { ok: false, error: 'Na tom mjestu nema slike.' };
    if (d.on && !cur.gid) { cur.gid = DriveApp.getFileById(cur.id).makeCopy(label + '.jpg', galerija_()).getId(); }
    if (!d.on && cur.gid) { trash_(cur.gid); delete cur.gid; }
    S[d.slot] = cur;
  } else if (d.akcija === 'brisi_sliku') {
    trash_(cur.id); trash_(cur.gid); delete S[d.slot];
  }
  setFields_(sh, row, { 'Slike': JSON.stringify(S) });
  return { ok: true, slike: S };
}

// Upiše vrijednosti u redak po nazivu stupca; stupci koji ne postoje dodaju se na kraj.
function setFields_(sh, row, fields) {
  let H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(fields).forEach(k => {
    let i = H.indexOf(k);
    if (i === -1) { i = H.length; sh.getRange(1, i + 1).setValue(k).setFontWeight('bold').setBackground('#FEE503'); H.push(k); }
    sh.getRange(row, i + 1).setValue(safe_(fields[k]));
  });
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
    return (typeof v === 'string' && /^[+0][\d\s\/().-]*$/.test(v)) ? "'" + v : safe_(v);
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

// Tekst koji počinje s = + - @ tablica bi shvatila kao formulu; ovako ostaje običan tekst.
function safe_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v) && v.charAt(0) !== "'") ? "'" + v : v;
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}


// ===== POTVRDA KLIJENTU =====
function potvrdaKlijentu_(data, nr, upit) {
  const g = k => { const key = Object.keys(data).find(x => x.indexOf('[' + k + ']') !== -1); return key ? String(data[key] || '') : ''; };
  const to = g('email');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(to)) return;
  const ime = [g('titula'), g('ime')].filter(String).join(' ');
  const box = 'font-family:Arial,sans-serif;font-size:15px;line-height:1.55;color:#111;max-width:560px';
  let subject, html;
  if (upit) {
    subject = 'Zaprimili smo Vaš upit – ' + SERVIS_NAZIV;
    html = '<div style="' + box + '"><p>Poštovani/a ' + esc_(ime) + ',</p><p>hvala na upitu. Zaprimili smo ga i javit ćemo se u najkraćem roku.</p>' +
      '<p style="background:#f4f4f0;padding:10px 14px;border-left:4px solid #FEE503">' + esc_(g('poruka')).replace(/\n/g, '<br>') + '</p>' +
      '<p>Za hitne slučajeve: <b>' + SERVIS_TEL + '</b> (Viber / WhatsApp)</p><p>Lijep pozdrav,<br>Saša Batinac<br>' + SERVIS_NAZIV + '</p></div>';
  } else {
    subject = 'Zaprimljen servisni nalog ' + nr + ' – ' + SERVIS_NAZIV;
    const ured = [g('uredaj'), g('tip_racunala'), g('laptop_model') || g('desktop_model')].filter(String).join(' · ');
    html = '<div style="' + box + '"><p>Poštovani/a ' + esc_(ime) + ',</p>' +
      '<p>zaprimili smo Vaš servisni nalog. Javit ćemo se s potvrdom termina.</p>' +
      '<p style="font-size:22px;font-weight:bold;margin:16px 0;background:#FEE503;display:inline-block;padding:6px 14px">' + nr + '</p>' +
      '<table style="border-collapse:collapse;font-size:14px">' +
      [['Uređaj', ured], ['Željeni datum', g('datum')], ['Predaja', g('predaja')], ['Opis kvara', g('opis_kvara')]].filter(r => r[1]).map(r => '<tr><td style="color:#666;padding:3px 14px 3px 0;vertical-align:top">' + r[0] + '</td><td style="padding:3px 0">' + esc_(r[1]) + '</td></tr>').join('') + '</table>' +
      '<p style="margin-top:18px"><b>Važno – backup podataka:</b> prije predaje uređaja napravite sigurnosnu kopiju svih važnih podataka (dokumenti, fotografije, lozinke, licence programa).</p>' +
      '<p>Status naloga možete provjeriti na <a href="' + WEB + 'status.html?nr=' + encodeURIComponent(nr) + '">' + WEB + 'status.html</a> uz broj naloga i broj mobitela.</p>' +
      '<p>Za hitne slučajeve: <b>' + SERVIS_TEL + '</b> (Viber / WhatsApp)</p><p>Lijep pozdrav,<br>Saša Batinac<br>' + SERVIS_NAZIV + '</p></div>';
  }
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: SERVIS_NAZIV, replyTo: NOTIFY_EMAIL });
}

// ===== JAVNA PROVJERA STATUSA (samo status, bez osobnih podataka) =====
function statusLookup_(nr, tel) {
  Utilities.sleep(600);
  nr = String(nr || '').trim().toUpperCase(); const t = String(tel || '').replace(/\D/g, '');
  if (!/^RN-\d{4}-\d{4}$/.test(nr) || t.length < 6) return { ok: false, error: 'Upišite broj naloga (npr. RN-2026-0001) i broj mobitela.' };
  const tb = table_(SHEET_NAME), H = tb.headers;
  const iN = H.indexOf('Broj naloga'), iT = H.indexOf('Kontakt broj mobitela'), iS = H.indexOf('Status'), iTe = H.indexOf('Termin'), iZ = H.indexOf('Završeno'), iJ = H.indexOf('Jamstvo do');
  const r = tb.rows.find(x => iN > -1 && String(x[iN]).toUpperCase() === nr && String(x[iT] || '').replace(/\D/g, '').slice(-6) === t.slice(-6));
  if (!r) return { ok: false, error: 'Nalog nije pronađen. Provjerite broj naloga i broj mobitela.' };
  const st = r[iS] || (r[iTe] ? 'Zakazano' : 'Novo');
  if (st === 'Arhivirano') return { ok: false, error: 'Nalog nije pronađen. Provjerite broj naloga i broj mobitela.' };
  const MAP = { 'Novo': ['zaprimljeno', 'Nalog je zaprimljen. Javit ćemo se s potvrdom termina.'], 'Zakazano': ['termin', 'Termin je dogovoren.'],
    'U radu': ['rad', 'Uređaj je na servisu, radimo na njemu.'], 'Pauzirano': ['rad', 'Uređaj je na servisu.'], 'Za obračun': ['rad', 'Servis je pri kraju.'],
    'Riješeno': ['gotovo', 'Servis je završen. Uređaj je spreman za preuzimanje.'], 'Odbijen': ['zatvoreno', 'Nalog je zatvoren. Za više informacija nazovite nas.'] };
  const m = MAP[st] || MAP['Novo'];
  return { ok: true, nr: nr, korak: m[0], poruka: m[1], termin: iTe > -1 ? r[iTe] : '', zavrseno: iZ > -1 ? r[iZ] : '', jamstvo: iJ > -1 ? r[iJ] : '' };
}

// ===== JUTARNJI SAŽETAK (svaki dan u 7:30) =====
function jutroAktivno_() { return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'jutarnjiSazetak'); }
function postaviJutarnjiMail_(on) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'jutarnjiSazetak').forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('jutarnjiSazetak').timeBased().everyDays(1).atHour(7).nearMinute(30).inTimezone(TZ).create();
}
function jutarnjiSazetak() {
  const tb = table_(SHEET_NAME), H = tb.headers, c = n => H.indexOf(n);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const rows = tb.rows.map(r => ({ nr: r[c('Broj naloga')] || '', ime: r[c('Ime i prezime')] || '', tel: r[c('Kontakt broj mobitela')] || '', ured: r[c('Vrsta uređaja')] || '',
    st: r[c('Status')] || (r[c('Termin')] ? 'Zakazano' : 'Novo'), term: r[c('Termin')] || '', t: r[c('vrijeme')] || '' }));
  const open = rows.filter(x => ['Riješeno', 'Odbijen', 'Arhivirano'].indexOf(x.st) === -1);
  const danas = open.filter(x => String(x.term).indexOf(today) === 0).sort((a, b) => a.term < b.term ? -1 : 1);
  const obr = open.filter(x => x.st === 'Za obračun' || x.st === 'Pauzirano');
  const novi = open.filter(x => x.st === 'Novo');
  const li = x => '<li><b>' + esc_(x.term ? String(x.term).slice(11) + ' ' : '') + esc_(x.ime) + '</b> · ' + esc_(x.ured) + ' · ' + esc_(x.tel) + ' <span style="color:#888">' + esc_(x.nr) + '</span></li>';
  const sec = (t, a) => '<h3 style="margin:18px 0 6px">' + t + ' (' + a.length + ')</h3>' + (a.length ? '<ul style="padding-left:18px;margin:0">' + a.map(li).join('') + '</ul>' : '<p style="color:#888;margin:0">Nema.</p>');
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#111;max-width:620px"><h2 style="margin:0;background:#FEE503;padding:8px 12px;display:inline-block">Dobro jutro! Pregled za ' + Utilities.formatDate(new Date(), TZ, 'd. M. yyyy.') + '</h2>' +
    '<p>Otvorenih naloga: <b>' + open.length + '</b></p>' + sec('Današnji termini', danas) + sec('Novi nalozi bez termina', novi) + sec('Na pauzi ili čeka obračun', obr) +
    '<p style="margin-top:20px"><a href="' + ADMIN_URL + '">Otvori admin</a></p></div>';
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: 'Jutarnji pregled: ' + danas.length + ' termina danas, ' + open.length + ' otvorenih naloga', htmlBody: html, name: SERVIS_NAZIV });
}
