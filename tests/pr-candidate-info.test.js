// หน้า PR Report ขั้นตอนที่ 3: การ์ดรายการต้องบอก ไลน์ / ประเภท (อะไหล่-สิ้นเปลือง-เครื่องมือช่าง) / ลงงบไหน
// และถ้ามาจากคำขอซื้อ ต้องโชว์รายละเอียดคำขอ (ขอโดย วันที่ เหตุผล หมายเหตุ) — เดิมบอกน้อยจนผู้ตรวจต้องย้อนไปดูหน้าขอซื้อ
const { html, grabFn, assert } = require('./_export-extract');

const src = ['isSharedLineName', 'fccLineOf', 'prItemTypeInfo', 'prBudgetInfo', 'prCandidateInfoHtml'].map(function(n) { return grabFn(n); }).join('\n');
const api = new Function('escHtml', src + '\nreturn { type: prItemTypeInfo, budget: prBudgetInfo, info: prCandidateInfoHtml };')(
  function(v) { return String(v).replace(/</g, '&lt;'); });

assert.strictEqual(api.type('Consumable').label, 'วัสดุสิ้นเปลือง');
assert.strictEqual(api.type('Tools').label, 'เครื่องมือช่าง');
assert.strictEqual(api.type('Office').label, 'อุปกรณ์สำนักงาน');
assert.strictEqual(api.type('Pneumatic').label, 'อะไหล่เครื่องจักร · หมวด Pneumatic', 'ของในสต็อกบอกหมวดอะไหล่ด้วย');
assert.strictEqual(api.type('-').label, 'อะไหล่เครื่องจักร');

assert.strictEqual(api.budget('ส่วนกลาง').label, '🏢 งบส่วนกลาง (ไม่ตัดงบไลน์)');
assert.strictEqual(api.budget('Lug&Screw').label, '💰 งบ Spare part ไลน์ Lug&Screw');
assert(/ยังไม่ระบุงบ/.test(api.budget('').label));

// เคสในภาพ: เครื่องเชื่อมจุดไฟ จากคำขอของ Nui ไลน์ส่วนกลาง
const card = api.info({ reason: 'REQ', line: 'ส่วนกลาง', category: 'Tools', request_id: 'REQ-1', requestedBy: 'Nui',
  requestedDate: '2026-11-07 09:28:36', requestReason: 'ใช้ซ่อมเครื่อง', requestRemark: 'จำเป็นในการ modify เครื่อง', isNewItemRequest: true });
['🏭 ไลน์: ส่วนกลาง', '🔧 เครื่องมือช่าง', '🏢 งบส่วนกลาง', 'รายละเอียดคำขอซื้อ', '#REQ-1', 'Nui', '2026-11-07 09:28', 'ใช้ซ่อมเครื่อง',
  'จำเป็นในการ modify เครื่อง', 'รายการใหม่'].forEach(function(t) { assert(card.includes(t), 'การ์ดต้องมี: ' + t); });
// รายการจาก Stock ไม่มีกล่องคำขอ แต่ยังบอกไลน์/ประเภท/งบ
const stockCard = api.info({ reason: 'BMIN', line: 'H9', category: 'Electrical' });
assert(stockCard.includes('💰 งบ Spare part ไลน์ H9') && !stockCard.includes('รายละเอียดคำขอซื้อ'));

// ฟอร์มขอซื้อ: รายการใหม่เลือกประเภทได้ และส่งไปเก็บในช่อง category
assert(html.includes('id="roNewItemCategory"'));
assert(html.includes("category: isNewItemMode ? ((document.getElementById('roNewItemCategory') || {}).value || '') : roSelectedItem.category,"));
// แก้ไขเพิ่มเติม: ย้ายไปลงงบส่วนกลางได้
assert(html.includes("['H9','Arc Chute','Coil Winding','Lug&Screw','ส่วนกลาง'].map(function(l){"));

console.log('pr-candidate-info: OK');
