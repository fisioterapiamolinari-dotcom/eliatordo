/**
 * Backend prenotazioni – Elia Tordo
 * Google Apps Script collegato a un Google Sheet (Estensioni → Apps Script).
 *
 * - Calendar: ogni prenotazione è un evento con SOLO il nome come titolo.
 * - Sheet "Prenotazioni": dati completi del cliente (cognome, telefono, email, token).
 * - Swap: i clienti si identificano con un token segreto (nel link della mail di conferma).
 * - n8n: riceve un webhook per ogni prenotazione / scambio / annullamento (per l'alarm).
 * - Promemoria: mail automatica il giorno prima (trigger giornaliero, vedi setup()).
 *
 * Proprietà dello script (Impostazioni progetto → Proprietà script):
 *   SITE_URL          https://TUOUTENTE.github.io/prenota-elia/
 *   N8N_WEBHOOK_URL   (facoltativa) URL del Webhook n8n
 *   N8N_SECRET        (facoltativa) stringa segreta, inviata nell'header X-Secret
 */

const CONFIG = {
  STUDIO_NAME: 'Elia Tordo',
  CALENDAR_ID: 'primary',        // oppure l'ID del calendario di Elia (condiviso con te in modifica)
  SHEET_NAME: 'Prenotazioni',
  TZ: 'Europe/Rome',             // imposta lo stesso fuso in appsscript.json
  DURATION_MIN: 45,
  STEP_MIN: 45,                  // distanza tra uno slot e il successivo
  DAYS_AHEAD: 42,                // finestra di prenotazione
  MIN_NOTICE_H: 12,              // preavviso minimo per prenotare
  MIN_SWAP_H: 24,                // niente scambi / annullamenti a meno di 24h
  // 0=dom 1=lun ... 6=sab. DA ADATTARE agli orari reali di Elia.
  ORARI: {
    1: [['09:00', '13:00'], ['15:00', '19:00']],
    2: [['09:00', '13:00'], ['15:00', '19:00']],
    3: [['09:00', '13:00'], ['15:00', '19:00']],
    4: [['09:00', '13:00'], ['15:00', '19:00']],
    5: [['09:00', '13:00']]
  }
};

const COL = { ID: 0, EVENT: 1, NOME: 2, COGNOME: 3, TEL: 4, EMAIL: 5, TOKEN: 6, SWAP: 7, DATA: 8, ORA: 9, STATO: 10, CREATO: 11, PROMEMORIA: 12 };
const HEADERS = ['id', 'eventId', 'nome', 'cognome', 'telefono', 'email', 'tokenHash', 'swapOk', 'data', 'ora', 'stato', 'creato', 'promemoria'];

/* ───────────── SETUP (da lanciare una volta a mano) ───────────── */
function setup() {
  sheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(18).inTimezone(CONFIG.TZ).create();
}

/* ───────────── ENDPOINT ───────────── */
function doGet(e) {
  try {
    const p = e.parameter || {};
    switch (p.action) {
      case 'getSlotsRange': return json_(getSlotsRange_(p.from, p.to));
      case 'getMine':       return json_(getMine_(p.token));
      case 'getSwapBoard':  return json_(getSwapBoard_(p.token));
      default:              return json_({ ok: false, error: 'azione_non_valida' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'errore_server' });
  }
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (b.action) {
        case 'book':    return json_(book_(b));
        case 'setSwap': return json_(setSwap_(b));
        case 'swap':    return json_(swap_(b));
        case 'cancel':  return json_(cancel_(b));
        default:        return json_({ ok: false, error: 'azione_non_valida' });
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'errore_server' });
  }
}

/* ───────────── DISPONIBILITÀ ───────────── */
function getSlotsRange_(from, to) {
  if (!isYmd_(from) || !isYmd_(to)) return { ok: false, error: 'dati_non_validi' };
  const today = parseStart_(todayYmd_(), '00:00');
  const maxDay = new Date(today.getFullYear(), today.getMonth(), today.getDate() + CONFIG.DAYS_AHEAD);
  let a = parseStart_(from, '00:00');
  let b = parseStart_(to, '00:00');
  if (a < today) a = today;
  if (b > maxDay) b = maxDay;
  const giorni = {};
  if (a > b) return { ok: true, giorni };
  const events = calendar_().getEvents(a, new Date(b.getFullYear(), b.getMonth(), b.getDate() + 1));
  for (let d = a; d <= b; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
    const ymd = Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd');
    giorni[ymd] = slotsForDay_(ymd, events);
  }
  return { ok: true, giorni };
}

function slotsForDay_(ymd, events) {
  const [y, m, d] = ymd.split('-').map(Number);
  const ranges = CONFIG.ORARI[new Date(y, m - 1, d).getDay()];
  if (!ranges) return [];
  const minStart = Date.now() + CONFIG.MIN_NOTICE_H * 3600000;
  const dur = CONFIG.DURATION_MIN * 60000;
  const out = [];
  ranges.forEach(r => {
    let t = parseStart_(ymd, r[0]).getTime();
    const end = parseStart_(ymd, r[1]).getTime();
    while (t + dur <= end) {
      const s = t, e = t + dur;
      const busy = events.some(ev => ev.getStartTime().getTime() < e && ev.getEndTime().getTime() > s);
      if (s >= minStart && !busy) out.push(Utilities.formatDate(new Date(s), CONFIG.TZ, 'HH:mm'));
      t += CONFIG.STEP_MIN * 60000;
    }
  });
  return out;
}

function eventsOfDay_(ymd) {
  const s = parseStart_(ymd, '00:00');
  return calendar_().getEvents(s, new Date(s.getFullYear(), s.getMonth(), s.getDate() + 1));
}

/* ───────────── PRENOTA ───────────── */
function book_(b) {
  if (b.website) return { ok: true, token: 'x' };            // honeypot: finto successo ai bot
  const nome = clean_(b.nome, 40), cognome = clean_(b.cognome, 60);
  const tel = clean_(b.telefono, 30), email = clean_(b.email, 100).toLowerCase();
  if (nome.length < 2 || cognome.length < 2 ||
      !/^[+0-9 ()\-]{6,}$/.test(tel) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !isYmd_(b.date) || !/^\d{2}:\d{2}$/.test(b.time || '')) {
    return { ok: false, error: 'dati_non_validi' };
  }
  const maxDay = new Date(Date.now() + CONFIG.DAYS_AHEAD * 86400000);
  if (parseStart_(b.date, b.time) > maxDay) return { ok: false, error: 'dati_non_validi' };

  if (slotsForDay_(b.date, eventsOfDay_(b.date)).indexOf(b.time) === -1) {
    return { ok: false, error: 'occupato' };
  }
  const now = Date.now();
  const haGia = rows_().some(x => x.r[COL.STATO] === 'attivo' &&
    String(x.r[COL.EMAIL]).toLowerCase() === email && startOf_(x.r).getTime() > now);
  if (haGia) return { ok: false, error: 'gia_prenotato' };

  const start = parseStart_(b.date, b.time);
  const end = new Date(start.getTime() + CONFIG.DURATION_MIN * 60000);
  const ev = calendar_().createEvent(nome, start, end);        // titolo = SOLO il nome

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  sheet_().appendRow([id, ev.getId(), nome, cognome, tel, email, hash_(token), false, b.date, b.time, 'attivo', new Date().toISOString(), '']);

  notify_({ type: 'booking', nome, cognome, telefono: tel, email, data: b.date, ora: b.time });
  mail_(email, 'Appuntamento confermato – ' + CONFIG.STUDIO_NAME,
    'Ciao ' + nome + ',\n\nil tuo appuntamento è confermato per ' + fmtLong_(start) + ' alle ' + b.time + '.\n\n' +
    'Per annullarlo, o scambiarlo con un altro cliente, usa questo link personale (non condividerlo):\n' +
    manageUrl_(token) + '\n\nTi scriverò un promemoria il giorno prima.\n\n' + CONFIG.STUDIO_NAME);
  return { ok: true, token };
}

/* ───────────── GESTIONE / SWAP ───────────── */
function getMine_(token) {
  const me = find_(token);
  if (!me) return { ok: false, error: 'token_non_valido' };
  const start = startOf_(me.r);
  if (start.getTime() < Date.now()) return { ok: true, booking: null };
  return { ok: true, booking: {
    nome: me.r[COL.NOME], date: txt_(me.r[COL.DATA], 'yyyy-MM-dd'), time: txt_(me.r[COL.ORA], 'HH:mm'),
    swapOk: isTrue_(me.r[COL.SWAP]), modificabile: start.getTime() - Date.now() > CONFIG.MIN_SWAP_H * 3600000
  } };
}

function getSwapBoard_(token) {
  const me = find_(token);
  if (!me) return { ok: false, error: 'token_non_valido' };
  const limite = Date.now() + CONFIG.MIN_SWAP_H * 3600000;
  const canSwap = startOf_(me.r).getTime() > limite;
  const items = !canSwap ? [] : rows_()
    .filter(x => x.r[COL.STATO] === 'attivo' && isTrue_(x.r[COL.SWAP]) &&
      x.r[COL.ID] !== me.r[COL.ID] && startOf_(x.r).getTime() > limite)
    .sort((p, q) => startOf_(p.r) - startOf_(q.r))
    .map(x => ({ id: x.r[COL.ID], date: txt_(x.r[COL.DATA], 'yyyy-MM-dd'), time: txt_(x.r[COL.ORA], 'HH:mm') }));
  return { ok: true, canSwap, items };       // niente nomi: solo data e ora
}

function setSwap_(b) {
  const me = find_(b.token);
  if (!me) return { ok: false, error: 'token_non_valido' };
  if (b.enabled && startOf_(me.r).getTime() - Date.now() <= CONFIG.MIN_SWAP_H * 3600000) {
    return { ok: false, error: 'troppo_vicino' };
  }
  sheet_().getRange(me.row, COL.SWAP + 1).setValue(!!b.enabled);
  return { ok: true };
}

function swap_(b) {
  const me = find_(b.token);
  if (!me) return { ok: false, error: 'token_non_valido' };
  const target = rows_().filter(x => x.r[COL.ID] === b.targetId && x.r[COL.STATO] === 'attivo')[0];
  if (!target || target.row === me.row || !isTrue_(target.r[COL.SWAP])) return { ok: false, error: 'non_disponibile' };

  const limite = Date.now() + CONFIG.MIN_SWAP_H * 3600000;
  if (startOf_(me.r).getTime() <= limite || startOf_(target.r).getTime() <= limite) {
    return { ok: false, error: 'troppo_vicino' };
  }
  const cal = calendar_();
  const evA = cal.getEventById(me.r[COL.EVENT]);
  const evB = cal.getEventById(target.r[COL.EVENT]);
  if (!evA || !evB) return { ok: false, error: 'non_disponibile' };

  const sA = evA.getStartTime(), eA = evA.getEndTime(), sB = evB.getStartTime(), eB = evB.getEndTime();
  evA.setTime(sB, eB);
  evB.setTime(sA, eA);

  const sh = sheet_();
  const fmtD = d => Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd');
  const fmtH = d => Utilities.formatDate(d, CONFIG.TZ, 'HH:mm');
  sh.getRange(me.row, COL.SWAP + 1, 1, 1).setValue(false);
  sh.getRange(target.row, COL.SWAP + 1, 1, 1).setValue(false);
  sh.getRange(me.row, COL.DATA + 1, 1, 2).setValues([[fmtD(sB), fmtH(sB)]]);
  sh.getRange(target.row, COL.DATA + 1, 1, 2).setValues([[fmtD(sA), fmtH(sA)]]);
  sh.getRange(me.row, COL.PROMEMORIA + 1).setValue('');
  sh.getRange(target.row, COL.PROMEMORIA + 1).setValue('');

  const info = (r, s) => ({ nome: r[COL.NOME], cognome: r[COL.COGNOME], telefono: r[COL.TEL], email: r[COL.EMAIL], data: fmtD(s), ora: fmtH(s) });
  notify_({ type: 'swap', a: info(me.r, sB), b: info(target.r, sA) });
  [[me.r, sB], [target.r, sA]].forEach(([r, s]) => mail_(String(r[COL.EMAIL]),
    'Scambio effettuato – ' + CONFIG.STUDIO_NAME,
    'Ciao ' + r[COL.NOME] + ',\n\nlo scambio è andato a buon fine: il tuo nuovo appuntamento è ' +
    fmtLong_(s) + ' alle ' + fmtH(s) + '.\n\n' + CONFIG.STUDIO_NAME));
  return { ok: true, date: fmtD(sB), time: fmtH(sB) };
}

function cancel_(b) {
  const me = find_(b.token);
  if (!me) return { ok: false, error: 'token_non_valido' };
  if (startOf_(me.r).getTime() - Date.now() <= CONFIG.MIN_SWAP_H * 3600000) return { ok: false, error: 'troppo_vicino' };
  try { const ev = calendar_().getEventById(me.r[COL.EVENT]); if (ev) ev.deleteEvent(); } catch (e) { console.error(e); }
  const sh = sheet_();
  sh.getRange(me.row, COL.SWAP + 1).setValue(false);
  sh.getRange(me.row, COL.STATO + 1).setValue('annullato');
  notify_({ type: 'cancel', nome: me.r[COL.NOME], cognome: me.r[COL.COGNOME], telefono: me.r[COL.TEL],
    email: me.r[COL.EMAIL], data: txt_(me.r[COL.DATA], 'yyyy-MM-dd'), ora: txt_(me.r[COL.ORA], 'HH:mm') });
  mail_(String(me.r[COL.EMAIL]), 'Appuntamento annullato – ' + CONFIG.STUDIO_NAME,
    'Ciao ' + me.r[COL.NOME] + ',\n\nil tuo appuntamento è stato annullato. Puoi prenotarne un altro quando vuoi:\n' +
    (prop_('SITE_URL') || '') + '\n\n' + CONFIG.STUDIO_NAME);
  return { ok: true };
}

/* ───────────── PROMEMORIA (trigger giornaliero) ───────────── */
function sendReminders() {
  const domani = Utilities.formatDate(new Date(Date.now() + 86400000), CONFIG.TZ, 'yyyy-MM-dd');
  const sh = sheet_();
  rows_().forEach(x => {
    if (x.r[COL.STATO] !== 'attivo' || x.r[COL.PROMEMORIA]) return;
    if (txt_(x.r[COL.DATA], 'yyyy-MM-dd') !== domani) return;
    const ora = txt_(x.r[COL.ORA], 'HH:mm');
    mail_(String(x.r[COL.EMAIL]), 'Promemoria: domani alle ' + ora + ' – ' + CONFIG.STUDIO_NAME,
      'Ciao ' + x.r[COL.NOME] + ',\n\nti ricordo l\'appuntamento di domani alle ' + ora + '.\n' +
      'Se non puoi venire, usa il link personale nella mail di conferma per annullare o scambiare.\n\n' + CONFIG.STUDIO_NAME);
    sh.getRange(x.row, COL.PROMEMORIA + 1).setValue(new Date().toISOString());
  });
}

/* ───────────── UTILITÀ ───────────── */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function calendar_() { return CONFIG.CALENDAR_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }
function isYmd_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s || ''); }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function todayYmd_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function parseStart_(ymd, hm) { const [y, m, d] = ymd.split('-').map(Number); const [h, mi] = hm.split(':').map(Number); return new Date(y, m - 1, d, h, mi); }
function startOf_(r) { return parseStart_(txt_(r[COL.DATA], 'yyyy-MM-dd'), txt_(r[COL.ORA], 'HH:mm')); }
function txt_(v, fmt) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, fmt) : String(v); }
function fmtLong_(d) { return Utilities.formatDate(d, CONFIG.TZ, 'dd/MM/yyyy'); }
function hash_(t) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join(''); }
function manageUrl_(token) { return (prop_('SITE_URL') || '') + '#gestisci=' + token; }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange('I:J').setNumberFormat('@');     // data e ora restano testo
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_() {
  const v = sheet_().getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ row: i + 2, r }));
}
function find_(token) {
  if (!token || typeof token !== 'string' || token.length < 32) return null;
  const h = hash_(token);
  return rows_().filter(x => x.r[COL.TOKEN] === h && x.r[COL.STATO] === 'attivo')[0] || null;
}

function mail_(to, subject, body) {
  try { MailApp.sendEmail({ to, subject, body, name: CONFIG.STUDIO_NAME }); } catch (e) { console.error('mail', e); }
}
function notify_(payload) {
  const url = prop_('N8N_WEBHOOK_URL');
  if (!url) return;
  payload.studio = CONFIG.STUDIO_NAME;
  payload.timestamp = new Date().toISOString();
  try {
    UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'X-Secret': prop_('N8N_SECRET') || '' }, payload: JSON.stringify(payload) });
  } catch (e) { console.error('n8n', e); }
}
