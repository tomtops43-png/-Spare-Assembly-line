const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── ฟอร์ม Manual ต้องกรอกได้หลายรายการ แล้วบันทึกครั้งเดียว ──────────────────
assert(htmlLf.includes('id="phManualRows"'), 'ต้องมีที่สำหรับแถวรายการ');
assert(htmlLf.includes('id="phManualAddRow"'), 'ต้องมีปุ่มเพิ่มรายการ');
assert(htmlLf.includes('id="phManualRowCount"') && htmlLf.includes('id="phManualGrandTotal"'),
  'ต้องสรุปจำนวนรายการและยอดรวมทั้งใบ');
assert(htmlLf.includes('id="phManualPartList"'), 'ต้องมี datalist ของอะไหล่รายไลน์');
// ช่องเดี่ยวของเดิมต้องถูกถอดออก ไม่งั้นเหลือสองทางกรอกที่ขัดกันเอง
['phManualPartName', 'phManualPartId', 'phPartNameDropdown', 'phPartNameClear'].forEach(function(id) {
  assert(!htmlLf.includes('id="' + id + '"'), 'ต้องไม่เหลือช่องกรอกแบบแถวเดียวของเดิม: ' + id);
});

// วาดแถวใหม่ทุกคีย์ = เคอร์เซอร์เด้งออกจากช่องที่กำลังพิมพ์ (บทเรียนเดียวกับค่าใช้จ่ายสิ้นเปลือง)
const recalc = grab(/^ {4}function phRecalcTotals\(\) \{[\s\S]*?\n {4}\}/m, 'phRecalcTotals');
assert(!/innerHTML/.test(recalc), 'ตอนพิมพ์ต้องอัปเดตแค่ตัวเลข ไม่วาดแถวใหม่');
assert(/rowsWrap\.addEventListener\('input', phOnRowInput\)/.test(htmlLf), 'ต้องผูกอีเวนต์แบบ delegate ที่กล่องแถว');

// ── เลือก Line ไหน ต้องขึ้นรายการอะไหล่ของไลน์นั้น ──────────────────────────
const sync = grab(/^ {4}function phSyncLinePartOptions\(\) \{[\s\S]*?\n {4}\}/m, 'phSyncLinePartOptions');
assert(sync.includes('phLoadLineParts(lineKey)'), 'ต้องโหลดอะไหล่ตามไลน์ที่เลือก');
assert(sync.indexOf('!== lineKey) return rows;') > -1,
  'ผู้ใช้เปลี่ยนไลน์ระหว่างรอโหลด ของที่มาช้าต้องไม่ทับลิสต์ของไลน์ปัจจุบัน');
const loadParts = grab(/^ {4}function phLoadLineParts\(lineKey\) \{[\s\S]*?\n {4}\}/m, 'phLoadLineParts');
assert(loadParts.includes("loadLineCache(key, key + '::ALL')"), 'ต้องใช้แคชรายไลน์ก้อนเดียวกับหน้า Stock');
assert(loadParts.includes("saveLineCache(key, rows, key + '::ALL')"), 'โหลดมาแล้วต้องเก็บแคชให้หน้าอื่นใช้ต่อ');
assert(/lineSelect\.addEventListener\('change', function\(\) \{ phSyncLinePartOptions\(\); \}\)/.test(htmlLf),
  'เปลี่ยน Line แล้วลิสต์ต้องเปลี่ยนตาม');
// เลือกอะไหล่แล้วเติมข้อมูลให้ แต่ห้ามทับของที่พิมพ์เองไว้
const autofill = grab(/^ {4}function phAutoFillRowFromMaster\(row, rowEl\) \{[\s\S]*?\n {4}\}/m, 'phAutoFillRowFromMaster');
assert(autofill.includes("if (String(row[key] || '').trim()) return;"), 'ห้ามทับค่าที่ผู้ใช้กรอกเอง');

// ── บันทึกครั้งเดียว + ถอยไปทีละแถวได้อย่างปลอดภัย ──────────────────────────
const submit = grab(/^ {4}function phSubmitManualForm\(\) \{[\s\S]*?\n {4}\}/m, 'phSubmitManualForm');
assert(submit.includes("action: 'addManualPurchaseHistoryBatch'"), 'ต้องยิง action แบบหลายรายการ');
assert(submit.includes('items: JSON.stringify(items)'), 'ต้องส่งรายการทั้งใบไปพร้อมกัน');
assert(/แถวที่ ' \+ \(i \+ 1\) \+ ': กรุณาระบุ Part Name/.test(submit), 'ต้องบอกว่าแถวไหนกรอกไม่ครบ');
assert(/แถวที่ ' \+ \(i \+ 1\) \+ ': Qty ต้องมากกว่า 0/.test(submit), 'ต้องกัน Qty ว่าง/ติดลบรายแถว');
// ถอยไปบันทึกทีละแถวได้เฉพาะตอนที่แน่ใจว่าแบ็กเอนด์ยังไม่รู้จัก action ใหม่เท่านั้น
// error อื่นห้ามถอย เพราะ batch อาจเขียนไปแล้วบางส่วน ยิงซ้ำจะได้รายการซ้ำ
assert(submit.includes('if (!phIsUnknownActionError(err)) throw err;'), 'error อื่นห้ามถอยไปยิงซ้ำ');
const unknown = grab(/^ {4}function phIsUnknownActionError\(err\) \{[\s\S]*?\n {4}\}/m, 'phIsUnknownActionError');
assert(unknown.includes('ต้องมี partName'), 'ต้องจับข้อความที่ backend รุ่นเก่าโยนออกมาก่อนเขียนชีท');
assert(htmlLf.includes('function phSaveManualRowsOneByOne(base, items, onProgress)'), 'ต้องมีทางบันทึกทีละแถว');
// backend รุ่นเก่าโยนข้อความนี้ตั้งแต่ก่อนเขียนอะไรลงชีท — ถอยแล้วจึงไม่เกิดรายการซ้ำ
assert(/throw new Error\('ต้องมี partName และ qty'\)/.test(backend), 'ข้อความที่ใช้ตรวจต้องยังอยู่ใน backend');

// ── Backend: action ใหม่ ────────────────────────────────────────────────────
assert(backend.includes("if (action === 'addManualPurchaseHistoryBatch') return respond(addManualPurchaseHistoryBatch(body), e);"),
  'ต้องต่อสาย action ใหม่ใน doPost');
const be = backend.match(/^function addManualPurchaseHistoryBatch\(payload\) \{[\s\S]*?\n\}/m);
assert(be, 'ต้องมีฟังก์ชัน addManualPurchaseHistoryBatch');
const beSrc = be[0];
assert(beSrc.includes("requireWarehouseWriter({ authToken: payload.authToken }, 'view_logs')"),
  'ต้องด่านสิทธิ์เท่ากับการบันทึกทีละรายการ');
assert(beSrc.includes('JSON.parse(items)'), 'รับ items ที่ส่งมาเป็นสตริงได้ (form-encoded)');
assert(beSrc.includes('items.length > 50'), 'ต้องกันบันทึกทีละมากเกินจน Apps Script timeout');
assert(beSrc.includes("source: 'Manual'"), 'ต้องลงเป็น source Manual เหมือนของเดิม');
assert(beSrc.includes('force_status: true'), 'ต้องเคารพสถานะที่ผู้ใช้เลือก');
assert(beSrc.includes('attachment_url: attachmentUrl'), 'ไฟล์แนบของใบต้องติดไปทุกแถว');
// แถวเดียวพังต้องไม่ล้มทั้งใบ แต่ถ้าไม่ผ่านสักแถวต้องโยน error ไม่ใช่รายงานว่าสำเร็จ
assert(beSrc.includes("errors.push('แถวที่ '"), 'เก็บข้อผิดพลาดรายแถวไว้รายงาน');
assert(beSrc.includes("if (!saved.length) throw new Error("), 'ไม่ผ่านสักแถวต้องถือว่าล้มเหลว');

// ── อัปโหลดไฟล์แนบต้องทนขึ้น (เดิมยิงครั้งเดียวจบ เจอ 404 ทีก็พังทั้งงาน) ────
const resilient = grab(/^ {4}function postFormResilient\(payload, attempts\) \{[\s\S]*?\n {4}\}/m, 'postFormResilient');
assert(resilient.includes('getManageApiCandidates()'), 'ต้องมี URL สำรองให้สลับ (/exec ↔ /dev)');
assert(resilient.includes('new URLSearchParams()'), 'ส่งแบบ form-encoded เหมือนท่ออัปโหลดรูปอะไหล่ที่ใช้ได้จริง');
assert(resilient.includes('return attempt(err);'), 'พลาดแล้วต้องยิงซ้ำ ไม่ใช่ล้มทันที');
assert(/var wait = idx < urls\.length \? 0 : 800;/.test(resilient), 'รอบถัดไปต้องหน่วงก่อน ไม่กระหน่ำยิง');
assert(htmlLf.includes('return postFormResilient(withAuthPayload({'), 'ไฟล์แนบต้องส่งผ่านท่อที่ยิงซ้ำได้');

// ── ย่อรูปต้องถึงเป้าขนาดจริง ไม่ใช่ย่อแค่ด้านกว้าง ─────────────────────────
const compress = grab(/^ {4}function compressImageToDataUrl\(file, options\) \{[\s\S]*?\n {4}\}/m, 'compressImageToDataUrl');
assert(compress.includes('out.length * 0.75 <= targetBytes'), 'ต้องวัดขนาดจริงหลังแปลง base64');
assert(compress.includes('if (quality > 0.6) quality -= 0.12; else dim = Math.round(dim * 0.8);'),
  'ยังใหญ่เกินต้องลดคุณภาพก่อน แล้วค่อยลดขนาดภาพ');
assert(/for \(var step = 0; step < 6; step \+= 1\)/.test(compress), 'ต้องมีเพดานรอบ ไม่วนไม่จบ');

console.log('Purchase history manual rows checks passed');

// ── พังกลางทางต้องไม่ทำให้ผู้ใช้กดซ้ำจนได้รายการซ้ำ ────────────────────────
const seq = grab(/^ {4}function phSaveManualRowsOneByOne\(base, items, onProgress\) \{[\s\S]*?\n {4}\}/m, 'phSaveManualRowsOneByOne');
assert(seq.includes("'บันทึกได้ ' + saved + '/' + items.length"), 'ต้องบอกว่าบันทึกไปแล้วกี่รายการก่อนหยุด');
assert(/กรุณาลบรายการที่บันทึกไปแล้วออกจากฟอร์มก่อนลองใหม่/.test(seq), 'ต้องบอกวิธีลองใหม่โดยไม่ให้ของซ้ำ');
assert(htmlLf.includes("if (rowErrors.length) { loadPurchaseHistory(true).catch(function() {}); return; }"),
  'มีแถวไม่ผ่านต้องคาฟอร์มไว้ให้อ่าน ไม่ปิดหนี');
assert(/ไม่ผ่าน ' \+ rowErrors\.length \+ ' แถว/.test(htmlLf), 'ต้องรายงานแถวที่ backend ตีกลับ');

console.log('Purchase history partial-failure checks passed');
