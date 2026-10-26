const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── ขอซื้อของที่ไม่มีในระบบ (เช่น แบตโน้ตบุ๊ก) ต้องแนบเอกสาร PDF ประกอบได้ ─────────────
assert(htmlLf.includes('id="roAttachmentFile" type="file" accept="image/*,application/pdf,.pdf"'), 'ช่องแนบไฟล์ต้องรับ PDF');
assert(/if \(isPdfFile\(file\)\) \{[\s\S]{0,400}file\.size > RO_ATTACH_PDF_MAX_BYTES/.test(htmlLf), 'PDF ต้องเช็คขนาดก่อนส่ง (บีบอัดแบบรูปไม่ได้)');
assert(/roAttachmentMimeType = 'application\/pdf';/.test(htmlLf));

const isRequestAttachmentDoc = new Function(grab(/^ {4}function isRequestAttachmentDoc\(url\) \{[\s\S]*?\n {4}\}/m, 'isRequestAttachmentDoc') + '\nreturn isRequestAttachmentDoc;')();
assert.strictEqual(isRequestAttachmentDoc('https://drive.google.com/file/d/abc123/view'), true);
assert.strictEqual(isRequestAttachmentDoc('https://drive.google.com/uc?export=view&id=abc123'), false, 'รูปเดิมต้องยังโชว์เป็นรูป');
assert.strictEqual(isRequestAttachmentDoc(''), false);
// PDF ห้ามถูกยัดเป็นรูปอะไหล่ใน PR หรือ <img> บนการ์ดคำขอ
assert(htmlLf.includes("var reqAttachImage = isRequestAttachmentDoc(r.attachment_url) ? '' : (r.attachment_url || '');"));
assert(htmlLf.includes('mainImage: (noKey && imageByNo[noKeyComposite]) || reqAttachImage,'));
assert(/var imgHtml = isRequestAttachmentDoc\(r\.attachment_url\)[\s\S]{0,400}📄 เปิดไฟล์แนบ \(PDF\)/.test(htmlLf));

// Backend รับ PDF และคืนลิงก์หน้าดูไฟล์ (ไม่ใช่ uc?export=view ที่เปิดเป็นรูป)
const upload = backend.slice(backend.indexOf('function uploadOrderRequestAttachmentToDrive(payload)'), backend.indexOf('function uploadRequestAttachment(payload)'));
assert(upload.includes("'application/pdf': 'pdf'"), 'backend ต้องรับ application/pdf');
assert(upload.includes("if (ext === 'pdf') return 'https://drive.google.com/file/d/' + file.getId() + '/view';"));
assert(backend.includes("attachmentUrl.match(/id=([^&]+)/) || attachmentUrl.match(/\\/file\\/d\\/([^/]+)/)"), 'ต้องดึง file id จากลิงก์ PDF ได้ด้วย');

// ── ตัวเลือก "ส่วนกลาง" สำหรับของที่ไม่ใช่อะไหล่ของไลน์ใด ─────────────────────────────
assert(/<select id="roNewItemLine"[\s\S]{0,600}<option value="ส่วนกลาง">/.test(htmlLf), 'รายการใหม่ต้องเลือกไลน์ "ส่วนกลาง" ได้');
assert(htmlLf.includes("line: isNewItemMode ? (document.getElementById('roNewItemLine').value || activeRealLine())"), 'คำขอรายการใหม่ต้องใช้ไลน์ที่เลือก');
const isSharedLineName = new Function(grab(/^ {4}function isSharedLineName\(value\) \{[\s\S]*?\n {4}\}/m, 'isSharedLineName') + '\nreturn isSharedLineName;')();
assert.strictEqual(isSharedLineName('ส่วนกลาง'), true);
assert.strictEqual(isSharedLineName(' ส่วนกลาง '), true);
assert.strictEqual(isSharedLineName('H9'), false);
// mapLine เดาทุกอย่างที่ไม่รู้จักเป็น H9 — คำขอส่วนกลางต้องไม่ถูกแปลงเป็น H9 แล้วไปกินงบ H9
assert(htmlLf.includes("var reqLine = isSharedLineName(r.line) ? 'ส่วนกลาง' : canonicalLineName(r.line || activeRealLine());"));
assert(htmlLf.includes('category: r.category || \'-\', line: reqLine,'));

console.log('Order request PDF + shared line checks passed');
