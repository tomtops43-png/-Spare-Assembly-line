// หน้าเว็บของระบบ PR → สั่งซื้อ → รับของ (GR): ตัวเลือก "รับตาม PR" ในฟอร์มรับเข้า + หน้า "ติดตาม PR"
const { html, grabFn, grabVar, assert } = require('./_export-extract');

// ---- ส่ง PR ที่เลือกไปกับการรับเข้า (ทั้ง POST และ JSONP) ----
const txnTransport = grabFn('callTransactionViaJsonp');
['prId: payload.prId', 'prLineNo: payload.prLineNo', 'noPrReason: payload.noPrReason'].forEach(function(f) {
  assert(txnTransport.includes(f), 'POST body ต้องมี ' + f);
});
["'prId=' + encodeURIComponent(requestPayload.prId)", "'prLineNo=' + encodeURIComponent(requestPayload.prLineNo)", "'noPrReason=' + encodeURIComponent(requestPayload.noPrReason)"].forEach(function(f) {
  assert(txnTransport.includes(f), 'JSONP query ต้องมี ' + f);
});

// ---- ฟอร์มรับเข้าทั้งสองจุดต้องตรวจ PR ก่อนส่ง และแนบค่าไปกับ payload ----
[['submitTransaction', 'txnPrPicker', 'txnPr'], ['submitQuickReceive', 'quickReceivePrPicker', 'quickPr']].forEach(function(t) {
  const body = grabFn(t[0]);
  const checkAt = body.indexOf("ptPickerValue('" + t[1] + "')");
  assert(checkAt > -1, t[0] + ' ต้องอ่านค่าจาก ' + t[1]);
  assert(checkAt < body.indexOf('callTransactionViaJsonp('), t[0] + ' ต้องตรวจ PR ก่อนส่งรายการ');
  assert(body.includes('if (!' + t[2] + '.ok)'), t[0] + ' ต้องหยุดถ้ายังไม่เลือก PR');
  ['prId: ' + t[2] + '.prId', 'prLineNo: ' + t[2] + '.prLineNo', 'noPrReason: ' + t[2] + '.noPrReason'].forEach(function(f) {
    assert(body.includes(f), t[0] + ' payload ต้องมี ' + f);
  });
  assert(body.includes('ptReceiptText(res)'), t[0] + ' ต้องบอกเลข GR / ยอดค้างหลังรับเข้า');
  assert(body.includes('ptAfterReceive()'), t[0] + ' ต้องรีเฟรชข้อมูลหน้าติดตาม PR');
});
assert(html.includes('<div id="txnPrPicker" class="hidden"></div>'), 'ฟอร์มรับเข้ามีช่องเลือก PR');
assert(html.includes('<div id="quickReceivePrPicker" class="hidden"></div>'), 'หน้าต่างรับเข้าด่วนมีช่องเลือก PR');
assert(grabFn('selectTxnItem').includes("ptRenderPicker('txnPrPicker', selected)"), 'เลือกอะไหล่แล้วโหลด PR ที่ค้างรับของตัวนั้น');
assert(grabFn('openQuickReceiveModal').includes("ptRenderPicker('quickReceivePrPicker', item)"), 'เปิดหน้าต่างรับเข้าแล้วโหลด PR ที่ค้างรับ');
assert(grabFn('ptRenderPicker').includes("action: 'getOpenPrLines', part_name: item.name || '', model: item.model || ''"), 'กรอง PR ตามอะไหล่ที่กำลังรับ');
assert(grabFn('ptRenderPicker').includes("isAdminUser()") && grabFn('ptRenderPicker').includes('__NO_PR__'), 'ตัวเลือก "รับโดยไม่มี PR" เฉพาะ Admin');

// ---- ตรรกะของตัวเลือก PR (รันจริงบน DOM จำลอง) ----
function makePicker(value, reason, ready) {
  return {
    getAttribute: function(k) { return k === 'data-pt-ready' && ready ? '1' : null; },
    querySelector: function(sel) { return sel === '[data-pt-select]' ? { value: value } : { value: reason || '' }; }
  };
}
function pickerApi(enabled, box) {
  return new Function('ptState', 'ptEl', grabFn('ptPickerValue') + '\nreturn ptPickerValue;')(
    { enabled: enabled }, function() { return box; });
}
assert.deepStrictEqual(pickerApi(false, null)('x'), { ok: true, prId: '', prLineNo: '', noPrReason: '' }, 'ยังไม่เปิดระบบ = รับเข้าแบบเดิม');
assert.strictEqual(pickerApi(true, makePicker('', '', false))('x').ok, false, 'ยังโหลด PR ไม่เสร็จ = ห้ามส่ง');
assert.strictEqual(pickerApi(true, makePicker('', '', true))('x').error, 'กรุณาเลือก PR ที่รับของเข้า');
assert.deepStrictEqual(pickerApi(true, makePicker('PR-2610-001|2', '', true))('x'), { ok: true, prId: 'PR-2610-001', prLineNo: '2', noPrReason: '' });
assert.strictEqual(pickerApi(true, makePicker('__NO_PR__', '  ', true))('x').ok, false, 'รับโดยไม่มี PR ต้องมีเหตุผล');
assert.deepStrictEqual(pickerApi(true, makePicker('__NO_PR__', 'ของแถม', true))('x'), { ok: true, prId: '', prLineNo: '', noPrReason: 'ของแถม' });

const receiptText = new Function(grabFn('ptFmtNum') + '\n' + grabFn('ptReceiptText') + '\nreturn ptReceiptText;')();
assert.strictEqual(receiptText({ goodsReceipt: { gr_id: 'GR-2610-0003', mode: 'PR', pr_id: 'PR-2610-001', qty_outstanding: 2, over_qty: 0 } }), ' • 🧾 GR-2610-0003 · PR-2610-001 ค้างรับอีก 2');
assert.strictEqual(receiptText({ goodsReceipt: { gr_id: 'GR-1', mode: 'PR', pr_id: 'PR-1', qty_outstanding: 0, over_qty: 3 } }), ' • 🧾 GR-1 · PR-1 รายการนี้รับครบแล้ว (รับเกิน 3)');
assert(receiptText({ goodsReceiptError: 'boom' }).includes('แจ้ง Admin'), 'บันทึกใบรับของพังต้องเตือน');
assert.strictEqual(receiptText({ purchaseHistorySync: {} }), '', 'ระบบเดิมไม่มีข้อความ GR');

// ---- หน้า "ติดตาม PR" ----
assert(html.includes('id="tabPrTracking"'), 'มีเมนูติดตาม PR ในกลุ่มจัดซื้อ');
assert(html.includes('<section id="prTrackingPage"'), 'มีหน้าติดตาม PR');
['open', 'list', 'gr'].forEach(function(t) { assert(html.includes('data-pt-tab="' + t + '"'), 'มีแท็บ ' + t); });
const loadFn = grabFn('ptLoad');
["action: 'getOpenPrLines'", "action: 'listPRs'", "action: 'getGrLog'"].forEach(function(a) { assert(loadFn.includes(a), 'หน้าติดตาม PR โหลด ' + a); });
assert(grabFn('ptOpenDetail').includes("action: 'getPRDetail'"), 'รายละเอียด PR โหลดจาก getPRDetail');
const actFn = grabFn('ptDetailAction');
["action: 'markPROrdered'", "action: 'cancelPR'", "action: 'closePrLine'"].forEach(function(a) { assert(actFn.includes(a), 'ปุ่มจัดการ PR เรียก ' + a); });
const startFn = grabFn('ptStartSystem');
assert(startFn.includes("action: 'startPrGrSystem'"), 'ปุ่มเริ่มระบบเรียก startPrGrSystem');
assert((startFn.match(/showConfirmDialog\(/g) || []).length === 2, 'เริ่มระบบต้องยืนยัน 2 ชั้น');
assert(grabFn('ptRenderBanner').includes('isAdminUser()'), 'ปุ่มเริ่มระบบเห็นเฉพาะ Admin');

// ทุกฟังก์ชันสลับหน้าต้องซ่อนหน้าติดตาม PR ด้วย (คู่กับ msHidePage เดิม) ไม่งั้นหน้าจะซ้อนกัน
const msCalls = (html.match(/^ *msHidePage\(\);$/gm) || []).length;
const ptCalls = (html.match(/^ *msHidePage\(\);\n *ptHidePage\(\);$/gm) || []).length;
assert(msCalls > 5 && msCalls === ptCalls, 'ptHidePage ต้องตามหลัง msHidePage ทุกจุด (' + ptCalls + '/' + msCalls + ')');
const tabStyles = grabFn('setTabStyles');
assert(tabStyles.includes("if (active === 'pr-tracking' && tabPrTrackingEl) tabPrTrackingEl.className = tabActiveClass;"), 'เมนูติดตาม PR ไฮไลต์ตอนเปิดหน้า');

// Tailwind 2.2.19 ที่โหลดอยู่ไม่มีสี emerald/orange/teal ฯลฯ — ต้องมี CSS กำหนดสีให้คลาสที่หน้านี้ใช้
['bg-emerald-600', 'bg-orange-100', 'bg-teal-100', 'text-rose-700'].forEach(function(cls) {
  assert(html.includes('#prTrackingPage .' + cls + ','), 'ต้องกำหนดสี .' + cls + ' ให้หน้าติดตาม PR');
});

console.log('PR tracking UI checks passed');
