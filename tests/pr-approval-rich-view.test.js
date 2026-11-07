// หน้าตรวจ+อนุมัติ PR: ผู้อนุมัติต้องเห็นภาพรวม (ใครขอ/ไลน์/ลงงบไหน/มูลค่ารวม) และเอกสารแนบโชว์ในหน้าเลย
// เดิมเห็นแค่ชื่อ+จำนวน และเอกสารแนบเป็นลิงก์ต้องกดเปิดทีละไฟล์
const { html, grabFn, assert } = require('./_export-extract');

const names = ['isSharedLineName', 'fccLineOf', 'prItemTypeInfo', 'prBudgetInfo', 'extractDriveFileId', 'buildDriveThumbnailUrl',
  'prApprMoney', 'prApprovalAttachmentHtml', 'prApprovalSummaryHtml'];
const api = new Function('escHtml', names.map(function(n) { return grabFn(n); }).join('\n') +
  '\nreturn { att: prApprovalAttachmentHtml, sum: prApprovalSummaryHtml };')(function(v) { return String(v).replace(/</g, '&lt;').replace(/"/g, '&quot;'); });

const ID = '1AbCdEfGhIjKlMnOpQrStUvWxYz012345';
// รูป → แสดงรูปเต็ม คลิกขยายได้
const img = api.att({ url: 'https://drive.google.com/file/d/' + ID + '/view', name: 'สกรีนช็อต 2026-11-07 112143.png' });
assert(img.includes('<img') && img.includes('thumbnail?id=' + ID + '&sz=w1600') && img.includes('data-appr-zoom='), 'รูปต้องแสดงในหน้า');
// PDF/เอกสาร → พรีวิว Drive ในกรอบ
const pdf = api.att({ url: 'https://drive.google.com/file/d/' + ID + '/view', name: 'ใบเสนอราคา.pdf' });
assert(pdf.includes('<iframe') && pdf.includes('https://drive.google.com/file/d/' + ID + '/preview'), 'PDF ต้องพรีวิวในหน้า');
assert(pdf.includes('เปิดแท็บใหม่'), 'ยังมีลิงก์เปิดไฟล์เต็ม');
// ชื่อไฟล์แปลกต้องไม่หลุด HTML
assert(!api.att({ url: 'https://drive.google.com/file/d/' + ID + '/view', name: '<b>x</b>.pdf' }).includes('<b>x</b>'));

// การ์ดสรุป: เคสในภาพ PR-2611-001 เครื่องเชื่อม 1 ชิ้น ไลน์ส่วนกลาง
const state = { header: { pr_id: 'PR-2611-001', created_by: 'admin', created_by_name: 'Admin', created_at: '2026-11-07 11:22:00', line: 'ส่วนกลาง' },
  lines: [{ line_no: 1, category: 'Tools', qty_approved: 1, unit_price: 1399 }, { line_no: 2, category: 'Consumable', qty_approved: 2, unit_price: 0 }] };
const sum = api.sum(state, {});
['PR-2611-001', 'รออนุมัติ', '🏢 งบส่วนกลาง', '🔧 เครื่องมือช่าง', '🧴 วัสดุสิ้นเปลือง', '฿1,399', 'Admin', '2026-11-07 11:22', 'ส่วนกลาง', '2 รายการ', '1 รายการยังไม่มีราคา']
  .forEach(function(t) { assert(sum.includes(t), 'การ์ดสรุปต้องมี: ' + t); });
// แก้ยอดอนุมัติ → มูลค่ารวมขยับตาม
assert(api.sum(state, { 1: 3 }).includes('฿4,197'));

// ตารางมีราคา/หน่วย + รวม และอัปเดตสดตอนแก้จำนวน
assert(html.includes('<th class="py-2 pr-2 text-right">ราคา/หน่วย</th>') && html.includes('data-appr-total="'));
assert(html.includes('renderPrApprovalSummary(); // ยอดรวมขยับทันทีที่แก้จำนวน'));
assert(html.includes("['prApprovalLines', 'prApprovalAttachments'].forEach(function(id) {"), 'คลิกรูปในเอกสารแนบเพื่อขยาย');

console.log('pr-approval-rich-view: OK');
