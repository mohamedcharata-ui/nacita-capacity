/**
 * NACITA Capacity Plan — online store for daily warehouse actuals.
 * Paste this whole file into Extensions → Apps Script of a Google Sheet,
 * change PASSCODE below, then Deploy → New deployment → Web app.
 */

// Leave empty ('') for no passcode. To require one later, put it here
// and set "requirePasscode": true in config.json on GitHub.
const PASSCODE = '';

const SHEET_NAME = 'DailyEntries';
const FIELDS = ['ibTrucks', 'obOrders', 'ppOcc', 'cbmOcc', 'fltOp', 'rtOp', 'hcPres', 'ibDocks', 'obDocks'];
const HEADERS = ['Date', 'Warehouse ID', 'IB Trucks', 'OB Orders', 'PP Occupied', 'CBM Occupied',
                 'FLT Operational', 'RT Operational', 'HC Present', 'IB Docks Used', 'OB Docks Used',
                 'Updated At', 'Updated By'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');   // keep dates as plain text (YYYY-MM-DD)
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).replace(/^'/, '');
}

/** GET → all daily entries (read-only, used by every page of the site). */
function doGet() {
  const values = getSheet_().getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0] || !r[1]) continue;
    const e = { date: dateText_(r[0]), whId: String(r[1]) };
    FIELDS.forEach(function (f, j) {
      const v = r[2 + j];
      if (v !== '' && v !== null) e[f] = Number(v);
    });
    e.updatedAt = r[11] instanceof Date ? r[11].toISOString() : String(r[11] || '');
    e.updatedBy = String(r[12] || '');
    entries.push(e);
  }
  return json_({ ok: true, entries: entries });
}

/** POST {passcode, editor, date, entries:[{whId, ibTrucks, ...}]} → create or update rows. */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (PASSCODE && String(body.passcode || '') !== PASSCODE) return json_({ ok: false, error: 'wrong passcode' });

  const date = String(body.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'bad date' });
  const editor = String(body.editor || 'unknown').slice(0, 60);
  const entries = Array.isArray(body.entries) ? body.entries : [];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    const rowOf = {};
    for (let i = 1; i < values.length; i++) rowOf[dateText_(values[i][0]) + '|' + values[i][1]] = i + 1;

    const now = new Date();
    let saved = 0;
    entries.forEach(function (en) {
      const wh = String(en.whId || '');
      if (!/^WH\d{2}$/.test(wh)) return;
      const row = ["'" + date, wh];
      FIELDS.forEach(function (f) {
        const x = en[f];
        const ok = x !== '' && x !== null && x !== undefined && !isNaN(Number(x)) && Number(x) >= 0;
        row.push(ok ? Number(x) : '');
      });
      row.push(now, editor);
      const key = date + '|' + wh;
      if (rowOf[key]) {
        sh.getRange(rowOf[key], 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
        rowOf[key] = sh.getLastRow();
      }
      saved++;
    });
    return json_({ ok: true, saved: saved, updatedAt: now.toISOString() });
  } finally {
    lock.releaseLock();
  }
}
