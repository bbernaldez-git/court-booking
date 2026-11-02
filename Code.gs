/** Basketball court booking backend. Bind this script to a Google Sheet (Extensions > Apps Script). */
const CONFIG = {
  SHEET: 'Bookings',
  FOLDER: 'Court Payment Proofs',
  COURT: 'Your Court Name',
  OPEN: 8,            // first slot starts 8:00 AM
  CLOSE: 22,          // last slot ends 10:00 PM
  RATE: 300,          // price per hour
  DAYS_AHEAD: 30,
  REFUND: [[48, 100], [24, 50]], // [hours before start, % of paid amount refunded]; less than the last = 0%
  PAY: 'GCash: 0917 000 0000 (Juan D.)\nMaya: 0918 000 0000 (Juan D.)'
};
const HEAD = ['Booking ID','Created','Date','Hours','Name','Phone','Email','Total','Payment Option','Amount To Pay','Balance Due','Proof','Status','Updated','Prev Status','Refund Amount','Refund Status','Refund Ref','Refund To'];
const HOLD = ['Pending','Confirmed','Cancel Requested']; // statuses that block a slot

function setup() { sheet_(); folder_(); }

function doGet(e) {
  const a = e.parameter.action;
  if (a === 'config') return out_({ ok: true, config: { court: CONFIG.COURT, open: CONFIG.OPEN, close: CONFIG.CLOSE, rate: CONFIG.RATE, days: CONFIG.DAYS_AHEAD, pay: CONFIG.PAY, refund: refundText_() } });
  if (a === 'slots') return out_({ ok: true, booked: booked_(e.parameter.date) });
  return out_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'book') return out_(book_(d));
    if (d.action === 'cancel') return out_(cancel_(d));
    if (!isAdmin_(d.password)) return out_({ ok: false, error: 'Wrong password' });
    if (d.action === 'list') return out_({ ok: true, bookings: list_() });
    if (d.action === 'update') return out_(update_(d.id, d.status, d.refund));
    if (d.action === 'proof') return out_(proof_(d.id));
    if (d.action === 'refunded') return out_(refunded_(d.id, d.ref));
    return out_({ ok: false, error: 'Unknown action' });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

function book_(d) {
  const tz = Session.getScriptTimeZone();
  const hours = [...new Set((d.hours || []).map(Number))].sort((a, b) => a - b);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) throw new Error('Invalid date');
  if (!hours.length || hours.length > 8 || hours.some(h => !Number.isInteger(h) || h < CONFIG.OPEN || h >= CONFIG.CLOSE)) throw new Error('Invalid time slots');
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const nowH = Number(Utilities.formatDate(new Date(), tz, 'H'));
  if (d.date < today || (d.date === today && hours.some(h => h <= nowH))) throw new Error('Those slots are in the past');
  const name = clean_(d.name), phone = clean_(d.phone);
  if (!name || !/^[0-9+\-\s()]{7,20}$/.test(phone)) throw new Error('Enter a valid name and phone number');
  if (!['half', 'full'].includes(d.payment)) throw new Error('Choose a payment option');
  if (!d.proof || !/^image\//.test(d.proofType || '')) throw new Error('Upload your payment screenshot');

  const total = hours.length * CONFIG.RATE;
  const due = d.payment === 'half' ? total / 2 : total;
  const bytes = Utilities.base64Decode(d.proof);
  if (bytes.length < 2000 || bytes.length > 2e6) throw new Error('Screenshot size is not valid');
  if (!(bytes[0] === -1 && bytes[1] === -40)) throw new Error('Invalid image file');
  const id = 'BK' + Utilities.formatDate(new Date(), tz, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  const file = folder_().createFile(Utilities.newBlob(bytes, 'image/jpeg', id + '.jpg')); // slow part, done BEFORE the lock

  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { file.setTrashed(true); throw new Error('Server is busy. Please try again.'); }
  try {
    const taken = booked_(d.date);                       // re-check inside the lock
    if (hours.some(h => taken.includes(h))) { file.setTrashed(true); throw new Error('Sorry, one of those slots was just booked. Please pick another.'); }
    const row = [id, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), d.date, hours.join(','), name, phone,
      clean_(d.email), total, d.payment === 'half' ? '50% downpayment' : 'Full payment', due, total - due, file.getUrl(), 'Pending', ''];
    const sh = sheet_(), r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row.map(String)]);
    SpreadsheetApp.flush();                              // make sure the write is visible before the next request reads
  } finally { lock.releaseLock(); }
  notify_('NEW BOOKING ' + id + '\n' + d.date + ' ' + hours.map(label_).join(', ') + '\n' + name + ' ' + phone +
    '\nPHP ' + due + ' of ' + total + ' (' + (d.payment === 'half' ? '50%' : 'full') + ')\nPlease verify payment.');
  return { ok: true, id: id, total: total, due: due };
}

function update_(id, status, refund) {
  if (!['Confirmed', 'Rejected', 'Pending', 'Cancelled'].includes(status)) throw new Error('Invalid status');
  const sh = sheet_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      if (HOLD.includes(status) && !HOLD.includes(rows[i][12])) {
        const taken = booked_(String(rows[i][2]));
        if (String(rows[i][3]).split(',').some(h => taken.includes(Number(h)))) throw new Error('Those slots were taken by another booking');
      }
      sh.getRange(i + 1, 13, 1, 2).setValues([[status, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]]);
      if (status === 'Cancelled') {
        const amt = Math.max(0, Number(refund) || 0);
        sh.getRange(i + 1, 16, 1, 2).setValues([[amt, amt > 0 ? 'Due' : 'None']]);
        if (rows[i][6]) { try { MailApp.sendEmail(rows[i][6], CONFIG.COURT + ' booking cancelled (' + id + ')', 'Hi ' + rows[i][4] + ',\n\nYour booking ' + id + ' has been cancelled.\n' + (amt > 0 ? 'Refund: PHP ' + amt + ' will be sent to ' + (rows[i][18] || rows[i][5]) + '.' : 'No refund applies.') + '\n'); } catch (e) {} }
      } else if (rows[i][12] === 'Cancel Requested') { sh.getRange(i + 1, 16, 1, 4).clearContent(); }
      if (status === 'Confirmed' && rows[i][6]) {
        try {
          MailApp.sendEmail(rows[i][6], CONFIG.COURT + ' booking confirmed (' + id + ')',
            'Hi ' + rows[i][4] + ',\n\nYour booking is confirmed.\nDate: ' + rows[i][2] + '\nTime: ' + rows[i][3].split(',').map(label_).join(', ') +
            '\nBalance due at the court: PHP ' + rows[i][10] + '\n\nSee you on the court!');
        } catch (e) {}
      }
      return { ok: true };
    }
  }
  throw new Error('Booking not found');
}

function cancel_(d) {
  const sh = sheet_(), rows = sh.getDataRange().getValues();
  const digits = x => String(x).replace(/\D/g, '').slice(-10);
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() === String(d.id || '').trim().toUpperCase() && digits(rows[i][5]) === digits(d.phone) && digits(d.phone).length >= 7) {
      if (!['Pending', 'Confirmed'].includes(rows[i][12])) throw new Error('This booking can no longer be cancelled online.');
      if (String(rows[i][2]) < Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')) throw new Error('This booking date has passed.');
      const start = Utilities.parseDate(rows[i][2] + ' ' + String(rows[i][3]).split(',')[0] + ':00', tz, 'yyyy-MM-dd H:mm');
      const hrs = (start - new Date()) / 36e5;
      let pct = 0; for (const [h, p] of CONFIG.REFUND) { if (hrs >= h) { pct = p; break; } }
      const refund = Math.round(Number(rows[i][9]) * pct / 100);
      sh.getRange(i + 1, 13, 1, 7).setValues([['Cancel Requested', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), rows[i][12], refund, 'Pending approval', '', clean_(d.refundTo) || rows[i][5]]]);
      notify_('CANCEL REQUEST ' + rows[i][0] + '\n' + rows[i][2] + ' ' + String(rows[i][3]).split(',').map(label_).join(', ') + '\n' + rows[i][4] + ' ' + rows[i][5] + '\nWas: ' + rows[i][12] + '\nSuggested refund: PHP ' + refund + ' (' + pct + '%)');
      return { ok: true, refund: refund, pct: pct };
    }
  }
  throw new Error('No booking matches that reference and phone number.');
}

function refundText_() {
  const r = CONFIG.REFUND;
  return r.map((x, i) => i === 0 ? 'Cancel ' + x[0] + '+ hours before your slot: ' + x[1] + '% refund' : 'Cancel ' + x[0] + ' to ' + r[i - 1][0] + ' hours before: ' + x[1] + '% refund').join('. ') + '. Less than ' + r[r.length - 1][0] + ' hours: no refund. Percentages apply to the amount you paid.';
}

function refunded_(id, ref) {
  const sh = sheet_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === id) {
    if (rows[i][16] !== 'Due') throw new Error('No refund is due for this booking');
    sh.getRange(i + 1, 17, 1, 2).setValues([['Refunded', clean_(ref)]]);
    return { ok: true };
  }
  throw new Error('Booking not found');
}

function proof_(id) {
  const rows = sheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === id) {
    const m = String(rows[i][11]).match(/[-\w]{25,}/);
    if (!m) throw new Error('No screenshot found');
    const b = DriveApp.getFileById(m[0]).getBlob();
    return { ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  }
  throw new Error('Booking not found');
}

function notify_(msg) {
  try {
    const p = PropertiesService.getScriptProperties();
    const key = p.getProperty('TEXTBEE_API_KEY'), dev = p.getProperty('TEXTBEE_DEVICE_ID'), to = p.getProperty('ADMIN_PHONE');
    if (!key || !dev || !to) return;
    UrlFetchApp.fetch('https://api.textbee.dev/api/v1/gateway/devices/' + dev + '/send-sms', {
      method: 'post', contentType: 'application/json', headers: { 'x-api-key': key },
      payload: JSON.stringify({ recipients: to.split(',').map(x => x.trim()), message: msg }), muteHttpExceptions: true
    });
  } catch (e) {}
}
function testSms() { notify_('Test SMS from your court booking app.'); }

function list_() {
  return sheet_().getDataRange().getValues().slice(1).reverse().map(r => Object.fromEntries(HEAD.map((h, i) => [h, String(r[i])]).filter(e => e[0] !== 'Proof'))); // Drive links are never sent to the browser
}

function booked_(date) {
  const out = [];
  sheet_().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[2]) === date && HOLD.includes(r[12])) String(r[3]).split(',').forEach(h => out.push(Number(h)));
  });
  return out;
}

function isAdmin_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!real && pw === real;
}
function label_(h) { h = Number(h); return (h % 12 || 12) + ':00 ' + (h < 12 ? 'AM' : 'PM'); }
function clean_(s) { s = String(s || '').trim().slice(0, 120); return /^[=+\-@]/.test(s) && !/^\+?[0-9]/.test(s) ? "'" + s : s; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET);
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1]) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  }
  return sh;
}
function folder_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.FOLDER);
  f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
