const SECRET = '';

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('postings') || ss.insertSheet('postings');
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const v = sheet().getDataRange().getValues(), rec = {};
  v.slice(1).forEach(r => { if (r[0]) rec[r[0]] = r[1]; });
  return out({ updated: new Date().toISOString(), rec: rec });
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.secret !== SECRET) return out({ ok: false, error: 'wrong password' });
  const s = sheet();
  s.clearContents();
  s.appendRow(['key', 'status']);
  const rows = Object.entries(d.rec);
  if (rows.length) s.getRange(2, 1, rows.length, 2).setValues(rows);
  return out({ ok: true });
}
