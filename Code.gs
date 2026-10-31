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
const HEAD = ['Booking ID','Created','Date','Hours','Name','Phone','Email','Total','Payment Option','Amount To Pay','Balance Due','Proof','Status','Updated'];
const HOLD = ['Pending','Confirmed']; // statuses that block a slot

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
    return { ok: true, id: id, total: total, due: due };
  } finally { lock.releaseLock(); }
}

function update_(id, status) {
  if (!['Confirmed', 'Rejected', 'Pending'].includes(status)) throw new Error('Invalid status');
  const sh = sheet_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
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
