// วางในชีตที่จะใช้เก็บข้อมูล: Extensions > Apps Script > วางไฟล์นี้ > Deploy > Web app
// (Execute as: Me, Who has access: Anyone) แล้วนำ URL ไปใส่ API ใน index.html
const PINS = { '1111': 'hk', '2222': 'store', '9999': 'admin' }; // เปลี่ยน PIN ก่อนใช้งานจริง
const COLLS = ['cfg', 'logs', 'sends', 'rets', 'cases', 'manual', 'closes'];

function sh(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(['id', 'updated', 'json']); }
  return s;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// โหลดข้อมูลทั้งหมดในครั้งเดียว
function doGet(e) {
  const role = PINS[e.parameter.pin];
  if (!role) return out({ err: 'pin' });
  const data = {};
  COLLS.forEach(c => {
    const s = sh(c), n = s.getLastRow();
    data[c] = n > 1 ? s.getRange(2, 3, n - 1, 1).getValues().map(r => JSON.parse(r[0])) : [];
  });
  return out({ role, data });
}

// บันทึกหลายรายการในครั้งเดียว (upsert ตาม id)
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (!PINS[b.pin]) return out({ err: 'pin' });
  const L = LockService.getScriptLock();
  L.waitLock(20000);
  try {
    const by = {};
    b.ops.forEach(o => (by[o.c] = by[o.c] || []).push(o.r));
    for (const c in by) {
      if (COLLS.indexOf(c) < 0) continue;
      const s = sh(c), n = s.getLastRow();
      const ids = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
      const add = [];
      by[c].forEach(r => {
        const row = [r.id, new Date(), JSON.stringify(r)], i = ids.indexOf(String(r.id));
        if (i < 0) { ids.push(String(r.id)); add.push(row); }
        else s.getRange(i + 2, 1, 1, 3).setValues([row]);
      });
      if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, 3).setValues(add);
    }
  } finally { L.releaseLock(); }
  return out({ ok: 1 });
}
