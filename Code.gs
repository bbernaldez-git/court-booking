/** Basketball court booking backend. Bind this script to a Google Sheet (Extensions > Apps Script). */
const CONFIG = {
  SHEET: 'Bookings',
  FOLDER: 'Court Payment Proofs',
  COURT: 'Your Court Name',
  OPEN: 8,            // first slot starts 8:00 AM
  CLOSE: 22,          // last slot ends 10:00 PM
  RATE: 300,          // price per hour
  DAYS_AHEAD: 30,
  PAY: 'GCash: 0917 000 0000 (Juan D.)\nMaya: 0918 000 0000 (Juan D.)'
};
const HEAD = ['Booking ID','Created','Date','Hours','Name','Phone','Email','Total','Payment Option','Amount To Pay','Balance Due','Proof','Status','Updated','Prev Status'];
const HOLD = ['Pending','Confirmed','Cancel Requested']; // statuses that block a slot

function setup() { sheet_(); folder_(); }

function doGet(e) {
  const a = e.parameter.action;
  if (a === 'config') return out_({ ok: true, config: { court: CONFIG.COURT, open: CONFIG.OPEN, close: CONFIG.CLOSE, rate: CONFIG.RATE, days: CONFIG.DAYS_AHEAD, pay: CONFIG.PAY } });
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
    if (d.action === 'update') return out_(update_(d.id, d.status));
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
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const taken = booked_(d.date);
    if (hours.some(h => taken.includes(h))) throw new Error('Sorry, one of those slots was just booked. Please pick another.');
    const id = 'BK' + Utilities.formatDate(new Date(), tz, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    const blob = Utilities.newBlob(Utilities.base64Decode(d.proof), d.proofType, id + '.jpg');
    if (blob.getBytes().length > 8e6) throw new Error('Screenshot is too large');
    const url = folder_().createFile(blob).getUrl();
    const sh = sheet_();
    const row = [id, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), d.date, hours.join(','), name, phone,
      clean_(d.email), total, d.payment === 'half' ? '50% downpayment' : 'Full payment', due, total - due, url, 'Pending', ''];
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row.map(String)]);
    notify_('NEW BOOKING ' + id + '\n' + d.date + ' ' + hours.map(label_).join(', ') + '\n' + name + ' ' + phone +
      '\nPHP ' + due + ' of ' + total + ' (' + (d.payment === 'half' ? '50%' : 'full') + ')\nPlease verify payment.');
    return { ok: true, id: id, total: total, due: due };
  } finally { lock.releaseLock(); }
}

function update_(id, status) {
  if (!['Confirmed', 'Rejected', 'Pending', 'Cancelled'].includes(status)) throw new Error('Invalid status');
  const sh = sheet_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      if (HOLD.includes(status) && !HOLD.includes(rows[i][12])) {
        const taken = booked_(String(rows[i][2]));
        if (String(rows[i][3]).split(',').some(h => taken.includes(Number(h)))) throw new Error('Those slots were taken by another booking');
      }
      sh.getRange(i + 1, 13, 1, 2).setValues([[status, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]]);
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
      sh.getRange(i + 1, 13, 1, 3).setValues([['Cancel Requested', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), rows[i][12]]]);
      notify_('CANCEL REQUEST ' + rows[i][0] + '\n' + rows[i][2] + ' ' + String(rows[i][3]).split(',').map(label_).join(', ') + '\n' + rows[i][4] + ' ' + rows[i][5] + '\nWas: ' + rows[i][12]);
      return { ok: true };
    }
  }
  throw new Error('No booking matches that reference and phone number.');
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
  return sheet_().getDataRange().getValues().slice(1).reverse().map(r => Object.fromEntries(HEAD.map((h, i) => [h, String(r[i])])));
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
    sh.getRange(1, HEAD.length).setValue(HEAD[HEAD.length - 1]).setFontWeight('bold');
  }
  return sh;
}
function folder_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.FOLDER);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
