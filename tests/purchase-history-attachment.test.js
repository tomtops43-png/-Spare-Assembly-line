const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ── ฟอร์ม "เพิ่มรายการ Manual" ต้องแนบบิลได้ ─────────────────────────────────
// ของที่ซื้อมามีบิลกระดาษเสมอ ถ้าแนบไม่ได้ก็ตรวจย้อนหลังไม่ได้ว่ายอดที่กรอกมาจากไหน
assert(htmlLf.includes('id="phManualPhotoFile"') && htmlLf.includes('id="phManualAttachFile"'),
  'ต้องมีทั้งปุ่มถ่ายรูปและปุ่มแนบไฟล์');
assert(/<input id="phManualPhotoFile"[^>]*capture="environment"/.test(htmlLf), 'ปุ่มถ่ายรูปต้องเปิดกล้องหลังบนมือถือ');
assert(/<input id="phManualAttachFile"[^>]*accept="image\/jpeg,image\/png,image\/webp,application\/pdf,\.pdf"/.test(htmlLf),
  'แนบไฟล์ต้องรับทั้งรูปและ PDF (ใบเสนอราคา/ใบส่งของมักเป็น PDF)');
assert(htmlLf.includes('id="phManualAttachUrl"') && htmlLf.includes('id="phManualAttachName"'));

// รูปจากมือถือ 4-8MB ส่งตรงไป Apps Script จะช้า/หลุด — ต้องย่อก่อน ส่วน PDF ห้ามย่อ (จะพัง)
assert(/var prepare = isImage \? resizeImageForAi\(file, 1600\) : readFileAsDataUrl\(file\);/.test(htmlLf));
assert(htmlLf.includes("action: 'uploadPurchaseHistoryAttachment'"));
assert(/month: phManualAttachMonth\(\),/.test(htmlLf), 'ต้องส่งเดือนไปด้วย ไฟล์จะได้ลงโฟลเดอร์เดือนที่ถูก');
assert(htmlLf.includes('var PH_ATTACH_MAX_BYTES = 8 * 1024 * 1024;') && htmlLf.includes('ไฟล์ใหญ่เกิน 8MB'));
assert(htmlLf.includes('รองรับเฉพาะไฟล์รูป (jpg/png/webp) และ PDF'));
// เลือกไฟล์เดิมซ้ำได้ ถ้าอัปโหลดรอบแรกพลาด (input ต้องถูกล้างค่า)
assert(/input\.value = ''; \/\/ เลือกไฟล์เดิมซ้ำได้/.test(htmlLf));

// เปิดฟอร์มใหม่ต้องล้างไฟล์แนบของรอบก่อน ไม่งั้นบิลของรายการเก่าติดไปกับรายการใหม่
assert(/document\.getElementById\('phManualPartId'\)\.value = '';\s*\n\s*phSetManualAttachment\('', ''\);/.test(htmlLf));
// ส่งไปกับรายการตอนบันทึก
assert(/attachment_url: document\.getElementById\('phManualAttachUrl'\)\.value,\s*\n\s*attachment_name: document\.getElementById\('phManualAttachName'\)\.value/.test(htmlLf));
// ตารางประวัติต้องเห็นว่ารายการไหนมีบิลแนบ
assert(htmlLf.includes('📎 ไฟล์แนบ'), 'ตาราง Purchase History ต้องมีลิงก์เปิดไฟล์แนบ');

// ── Backend: คอลัมน์ใหม่ + Drive แยกโฟลเดอร์ ─────────────────────────────────
assert(backend.includes("'Price Status', 'Created By', 'Request Period', 'Attachment URL', 'Attachment Name'];"),
  'คอลัมน์ไฟล์แนบต้องต่อท้าย ตำแหน่งคอลัมน์เดิมจะได้ไม่ขยับ');
// แถวเก่าที่ยังไม่มีไฟล์แนบต้อง migrate เป็นค่าว่าง ไม่ใช่ undefined (Sheets จะเขียนไม่ผ่าน)
assert(/getPurchaseHistoryCell\(row, oldMap, \['Attachment URL', 'attachment_url'\], ''\)/.test(backend));
// เพิ่มแถวใหม่ = เขียนไฟล์แนบ / อัปเดตแถวเดิม = ไม่ส่งมาก็ต้องคงของเดิมไว้
assert(/payload\.attachment_url \|\| payload\.attachmentUrl \|\| ''/.test(backend));
assert(/payload\.attachment_url !== undefined \? payload\.attachment_url : \(existing\[31\] \|\| ''\)/.test(backend));
assert(/attachment_url: String\(row\[31\] \|\| ''\), attachment_name: String\(row\[32\] \|\| ''\)/.test(backend));
assert(/function addManualPurchaseHistory\(payload\)[\s\S]{0,1400}attachment_url: String\(payload\.attachment_url \|\| payload\.attachmentUrl \|\| ''\)\.trim\(\)/.test(backend));

// ไฟล์ลง Drive คนละโฟลเดอร์กับงานอื่น แยกย่อยรายเดือนไว้ไล่หาตอนปิดบัญชี
assert(/getOrCreateChildFolder\(getOrCreateChildFolder\(root, 'purchase-history'\), month\)/.test(backend));
assert(backend.indexOf(".test(month)) month = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');") > -1,
  'เดือนที่ส่งมาผิดรูปแบบต้อง fallback เป็นเดือนปัจจุบัน ไม่ใช่สร้างโฟลเดอร์ชื่อมั่ว');
assert(/'image\/jpeg': 'jpg', 'image\/png': 'png', 'image\/webp': 'webp', 'application\/pdf': 'pdf'/.test(backend));
assert(backend.includes("throw new Error('ไฟล์แนบรองรับเฉพาะ jpg, png, webp, pdf');"));
// PDF ต้องได้ลิงก์ /view (เปิดอ่านในเบราว์เซอร์ได้) ส่วนรูปต้องเป็น uc เพื่อให้ <img> โหลดตรงๆ
assert(/mimeType === 'application\/pdf'[\s\S]{0,220}drive\.google\.com\/file\/d\/[\s\S]{0,220}uc\?export=view/.test(backend));
assert(backend.includes("if (action === 'uploadPurchaseHistoryAttachment') return respond(uploadPurchaseHistoryAttachment(body), e);"));

console.log('Purchase history attachment checks passed');
