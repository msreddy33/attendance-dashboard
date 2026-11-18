/**
 * Attendance posting dashboard — Apps Script backend
 *
 * SETUP (one time)
 * 1. Open the Google Sheet -> Extensions -> Apps Script, and paste this file as Code.gs.
 *    (If the script is NOT bound to the sheet, add a Script property SHEET_ID = <sheet id>.)
 * 2. Project Settings -> Script properties -> Add:  ADMIN_PASSWORD = <your password>
 * 3. Deploy -> Manage deployments -> pencil icon -> Version: "New version" -> Deploy
 *    Execute as: Me   |   Who has access: Anyone
 *    (editing the existing deployment keeps the same /exec URL)
 *
 * API
 *   GET  ?action=get                      -> public, read-only: { ok, rec, updated }
 *   POST {action:'login', password}       -> { ok } (checks the password)
 *   POST {action:'set', password, changes:{ "YYYY-MM-DD|A1|0": 0|1|2 }}
 *        1 = Posted, 2 = Not posted, 0 = clear (Pending). Requires the password.
 */

const SHEET_NAME = 'Marks';
const HEADERS = ['Key', 'Date', 'Section', 'Period', 'Status', 'Updated'];
const PERIODS = ['P1\u2013P2', 'P3\u2013P4', 'P5\u2013P6', 'P7\u2013P8'];
const STATUS_TEXT = { 1: 'Posted', 2: 'Not posted' };
const STATUS_NUM = { 'Posted': 1, 'Not posted': 2 };
const KEY_RE = /^\d{4}-\d{2}-\d{2}\|[A-Za-z0-9]+\|[0-3]$/;

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!passwordOk_(body.password)) {
      Utilities.sleep(800); // slows down password guessing
      return json_({ ok: false, error: 'unauthorized' });
    }
    if (body.action === 'login') return json_({ ok: true });
    if (body.action === 'set') return json_(applyChanges_(body.changes || {}));
    return json_({ ok: false, error: 'bad_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function passwordOk_(p) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real || typeof p !== 'string' || p.length !== real.length) return false;
  let diff = 0;
  for (let i = 0; i < p.length; i++) diff |= p.charCodeAt(i) ^ real.charCodeAt(i);
  return diff === 0;
}

function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const rec = {};
  if (n > 0) {
    sh.getRange(2, 1, n, 5).getValues().forEach(function (r) {
      const s = STATUS_NUM[r[4]];
      if (r[0] && s) rec[String(r[0])] = s;
    });
  }
  const updated = PropertiesService.getScriptProperties().getProperty('UPDATED') || '';
  return { ok: true, rec: rec, updated: updated };
}

function applyChanges_(changes) {
  const keys = Object.keys(changes);
  if (keys.length > 1000) return { ok: false, error: 'too_many_changes' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const now = new Date();
    const last = sh.getLastRow();
    const map = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 6).getValues().forEach(function (r) {
        if (r[0] && STATUS_NUM[r[4]]) map[String(r[0])] = { s: STATUS_NUM[r[4]], u: r[5] };
      });
    }

    keys.forEach(function (k) {
      if (!KEY_RE.test(k)) return;
      const v = Number(changes[k]);
      if (v === 1 || v === 2) {
        if (!map[k] || map[k].s !== v) map[k] = { s: v, u: now };
      } else {
        delete map[k];
      }
    });

    const rows = Object.keys(map).sort().map(function (k) {
      const p = k.split('|');
      return [k, p[0], p[1], PERIODS[Number(p[2])], STATUS_TEXT[map[k].s], map[k].u];
    });

    if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();
    if (rows.length) {
      sh.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // keep dates as plain text
      sh.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    const updated = now.toISOString();
    PropertiesService.getScriptProperties().setProperty('UPDATED', updated);
    return { ok: true, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
