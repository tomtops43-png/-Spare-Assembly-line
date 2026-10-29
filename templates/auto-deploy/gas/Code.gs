// ตัวอย่าง backend — แทนที่ด้วยโค้ดของเว็บคุณได้เลย
// หน้าเว็บเรียก:  GET  <URL>/exec?action=ping
//                POST <URL>/exec  body = JSON.stringify({ action: 'ping' })
// (POST ให้ส่ง Content-Type: text/plain เพื่อไม่ให้เบราว์เซอร์ยิง preflight ที่ GAS ไม่รองรับ)

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(body);
}

function handle_(req) {
  var result;
  try {
    switch (req.action) {
      case 'ping':
        result = { ok: true, time: new Date().toISOString() };
        break;
      default:
        result = { ok: false, error: 'ไม่รู้จัก action: ' + req.action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
