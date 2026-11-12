// Library backend (Google Apps Script). Paste in Extensions > Apps Script of a Google Sheet.
const PIN = '1234';               // CHANGE THIS to your staff PIN
const SHOW_NAMES_PUBLIC = false;  // false = only staff can see who has the book
const TAB = 'Transactions';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(TAB);
  if (!s) {
    s = ss.insertSheet(TAB);
    s.appendRow(['Time', 'Acc No', 'Reg No', 'Name', 'Action', 'Due Date', 'Phone']);
    s.getRange('B:C').setNumberFormat('@'); s.getRange('G:G').setNumberFormat('@');
  }
  return s;
}
function current_() {
  const v = sheet_().getDataRange().getValues(); v.shift();
  const m = {};
  v.forEach(r => {
    const acc = String(r[1]);
    if (r[4] === 'ISSUE') m[acc] = { reg: String(r[2]), name: r[3], phone: String(r[6]), due: Utilities.formatDate(new Date(r[5]), 'Asia/Kolkata', 'yyyy-MM-dd') };
    else delete m[acc];
  });
  return m;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function pub_(m) { const o = {}; for (const k in m) o[k] = { due: m[k].due }; return o; }

function books_() {
  const s = SpreadsheetApp.getActive().getSheetByName('Books');
  if (!s) return [];
  const v = s.getDataRange().getValues(); v.shift();
  return v.filter(r => String(r[2]).trim()).map(r => r.slice(0, 8).map(c => String(c).trim()));
}

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'books') return out_({ ok: true, books: books_() });
  const m = current_();
  return out_({ ok: true, issued: SHOW_NAMES_PUBLIC ? m : pub_(m) });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.pin !== PIN) return out_({ ok: false, error: 'Wrong PIN' });
  if (d.action === 'status') return out_({ ok: true, issued: current_() });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const cur = current_(), acc = String(d.acc);
    if (d.action === 'issue') {
      if (cur[acc]) return out_({ ok: false, error: 'This book is already issued' });
      const due = new Date(); due.setDate(due.getDate() + (+d.days || 14));
      sheet_().appendRow([new Date(), acc, String(d.reg), d.name, 'ISSUE', due, String(d.phone)]);
    } else if (d.action === 'return') {
      if (!cur[acc]) return out_({ ok: false, error: 'This book is not issued' });
      sheet_().appendRow([new Date(), acc, cur[acc].reg, cur[acc].name, 'RETURN', '', cur[acc].phone]);
    }
    return out_({ ok: true, issued: current_() });
  } finally { lock.releaseLock(); }
}
