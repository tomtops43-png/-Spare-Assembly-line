// ➕ เพิ่มรายการนอกระบบในหน้า PR Report (ขั้นตอนที่ 3)
// เดิมของที่ไม่มีในสต็อกต้องไปหน้า "ขอซื้อ" ก่อน แล้วค่อยมาเปิด PR โหมด "PR ขอซื้ออะไหล่" — ต้องเข้า 2 หน้า
// ตอนนี้กรอกในหน้า PR ได้เลย ระบบสร้างคำขอซื้อเบื้องหลัง (ได้ request_id ไว้ผูกกับเลข PR / ติดตาม / รับของ)
const { html, grabFn, assert } = require('./_export-extract');

// ---- UI ----
assert(html.includes('id="prAddManualBtn"'), 'มีปุ่มเพิ่มรายการนอกระบบในแถบเครื่องมือขั้นตอนที่ 3');
assert(html.includes('id="prManualModal"') && html.includes('id="prManualForm"'), 'มีหน้าต่างกรอกรายการ');
['prManualName', 'prManualModel', 'prManualBrand', 'prManualLine', 'prManualCategory', 'prManualQty', 'prManualUnit', 'prManualPrice', 'prManualPriority', 'prManualReason']
  .forEach(function(id) { assert(html.includes('id="' + id + '"'), 'มีช่อง ' + id); });
assert(html.includes('<option value="NEW">'), 'กรองเหตุผล "นอกระบบ" ได้');
assert(html.includes("NEW: ['🆕 นอกระบบ'"), 'มีป้าย 🆕 นอกระบบ บนการ์ด');

// ---- buildManualPrRequestPayload ----
const build = new Function(grabFn('buildManualPrRequestPayload') + '\nreturn buildManualPrRequestPayload;')();
const ok = build({ name: ' Proximity Sensor ', model: 'E2E-X5', brand: 'Omron', line: 'Arc Chute', category: 'Spare Part',
  qty: '2', unit: '', price: '1500', priority: 'Urgent', reason: 'ใช้ซ่อมเครื่อง' });
assert.strictEqual(ok.action, 'createOrderRequest');
assert.strictEqual(ok.item_id, '', 'ไม่มีเลขอะไหล่ = รายการใหม่');
assert.strictEqual(ok.item_name, 'Proximity Sensor');
assert.strictEqual(ok.request_qty, 2);
assert.strictEqual(ok.unit, 'PCS', 'หน่วยว่าง = PCS');
assert.strictEqual(ok.unit_price, 1500);
assert.strictEqual(ok.priority, 'Urgent');
assert.strictEqual(ok.line, 'Arc Chute');
assert.strictEqual(build({ name: 'x', qty: 1, reason: 'r', price: '' }).unit_price, '', 'ไม่ใส่ราคา = TBC');
assert.strictEqual(build({ name: 'x', qty: 1, reason: 'r', priority: 'อะไรก็ได้' }).priority, 'Normal');
assert(build({ name: '  ', qty: 1, reason: 'r' }).error, 'ต้องมีชื่อ');
assert(build({ name: 'x', qty: 0, reason: 'r' }).error, 'จำนวนต้อง > 0');
assert(build({ name: 'x', qty: 'abc', reason: 'r' }).error, 'จำนวนต้องเป็นตัวเลข');
assert(build({ name: 'x', qty: 1, reason: '' }).error, 'ต้องมีเหตุผล');
assert(build({ name: 'x', qty: 1, reason: 'r', price: '-5' }).error, 'ราคาติดลบไม่ได้');

// ---- orderRequestToPrCandidate: ใช้ร่วมกับโหมด "PR ขอซื้ออะไหล่" ----
const toCand = new Function('isSharedLineName', 'canonicalLineName', 'activeRealLine', 'isRequestAttachmentDoc',
  grabFn('orderRequestToPrCandidate') + '\nreturn orderRequestToPrCandidate;')(
  function(l) { return l === 'ส่วนกลาง'; }, function(l) { return l; }, function() { return 'H9'; }, function() { return false; });
const c = toCand(Object.assign({}, ok, { request_id: 'REQ-77', status: 'Pending' }), {});
assert.strictEqual(c.request_id, 'REQ-77', 'การ์ดมี request_id → ตอนส่ง PR ผูกคำขอกับเลข PR ได้');
assert.strictEqual(c.isNewItemRequest, true);
assert.strictEqual(c.line, 'Arc Chute');
assert.strictEqual(c.suggestedQty, 2);
assert.strictEqual(c.unitPrice, 1500);
assert.strictEqual(c.critical, 'Critical', 'ด่วน = Critical');
assert.strictEqual(c.selected, true);
assert(html.includes('return rows.map(function(r) { return orderRequestToPrCandidate(r, imageByNo); });'), 'โหมด PR ขอซื้อใช้ตัวแปลงเดียวกัน');

// ---- ผูกการทำงาน ----
const i = html.indexOf("if (prManualForm) prManualForm.addEventListener('submit'");
assert(i > -1, 'มีตัวจัดการกดเพิ่ม');
const handler = html.slice(i, i + 3500);
assert(handler.includes('payload.client_uid = prManualUid;'), 'กันคำขอซ้ำเวลากดใหม่');
assert(handler.includes('prCandidateState.unshift(cand);'), 'ใส่รายการเข้า PR ทันที');
assert(handler.includes('renderPrCandidates();'));
assert(handler.includes("cand.reason = 'REQ, NEW';"));
// ตอนส่ง PR บรรทัดที่มี request_id ถูกส่งไปให้ backend ผูก (กลไกเดิม)
assert(html.includes("payload.source_request_ids = sourceRequestIds.join(',');"));

console.log('pr-manual-out-of-system-item: OK');
