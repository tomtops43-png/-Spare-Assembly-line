const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// บั๊กจริง: ผู้ใช้เลือก Coil Size ใหม่ในฟอร์มแล้ว (ทั้งฟอร์ม "เพิ่มรายการใหม่" #mCoilSize
// และ "แก้ไขรายการ" #eCoilSize) กด save ผ่านไป แต่กลับมาเปิดดูใหม่ค่าไม่เปลี่ยน — เพราะ
// callManageAction() ที่ทั้งสองฟอร์มใช้ร่วมกัน ไม่เคยส่ง coil_size ออกไปเลยทั้งใน
// requestPayload (POST body) และ query string (GET/JSONP) แม้ payload ที่รับเข้ามาจะมีค่าอยู่แล้ว
// เว็บนี้รันบน GitHub Pages ส่วน backend อยู่คนละ origin (script.google.com) — ต่าง origin
// ทำให้ shouldUseJsonpTransportForUrl() คืนค่า true และ JSONP/GET (ผ่าน doGet) กลายเป็น
// transport หลักที่ใช้จริง ไม่ใช่ POST — ฝั่ง doGet ของ Backend.gs ก็ขาด coil_size เช่นกัน

const callManageActionBody = (function() {
  const start = html.indexOf('function callManageAction(action, payload)');
  const end = html.indexOf('\n    function callTransactionViaJsonp', start);
  assert(start > -1 && end > start, 'can slice callManageAction body');
  return html.slice(start, end);
})();

// ---- requestPayload (ใช้เป็น POST body ผ่าน doPost) ----
assert(/coil_size: payload\.coil_size \|\| ''/.test(callManageActionBody), 'requestPayload forwards payload.coil_size');

// ---- query string (ใช้เป็น GET/JSONP ผ่าน doGet — transport จริงเมื่อรันคนละ origin) ----
assert(/'coil_size=' \+ encodeURIComponent\(requestPayload\.coil_size\)/.test(callManageActionBody), 'GET/JSONP query string includes coil_size');

// ---- Backend.gs: ทั้ง doGet และ doPost ต้องส่ง coil_size เข้า upsertMainItem ----
const doGetUpsertBlock = (function() {
  const start = backend.indexOf("if (action === 'upsertItem') {");
  const end = backend.indexOf('\n    }', start);
  assert(start > -1 && end > start, 'can slice doGet upsertItem block');
  return backend.slice(start, end);
})();
assert(doGetUpsertBlock.includes('coil_size: e.parameter.coil_size,'), 'doGet (GET/JSONP) passes e.parameter.coil_size through to upsertMainItem');

const doPostUpsertBlock = (function() {
  const start = backend.indexOf("if (action === 'upsertItem') {", backend.indexOf("if (action === 'upsertItem') {") + 1);
  const end = backend.indexOf('\n    }', start);
  assert(start > -1 && end > start, 'can slice doPost upsertItem block');
  return backend.slice(start, end);
})();
assert(doPostUpsertBlock.includes('coil_size: body.coil_size || body.machine_model,'), 'doPost still passes body.coil_size through (regression guard on the already-working path)');

console.log('coil-size-save-transport: all assertions passed');
