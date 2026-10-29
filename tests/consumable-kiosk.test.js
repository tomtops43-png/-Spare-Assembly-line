const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(src, re, label) {
  const m = src.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── Backend: เบิกทั้งชุด ของไม่พอแม้ตัวเดียวต้องไม่ตัดอะไรเลย ─────────────────────
const plan = new Function(grab(backend, /^function planConsumableWithdraw\(itemsById, requested\) \{[\s\S]*?\n\}/m, 'planConsumableWithdraw') + '\nreturn planConsumableWithdraw;')();
const stock = {
  CS0001: { item_id: 'CS0001', name: 'ฟิตติ้ง PC 6-01', qty: 10, unit: 'ตัว', active: true },
  CS0002: { item_id: 'CS0002', name: 'น็อต M8', qty: 3, unit: 'ตัว', active: true },
  CS0009: { item_id: 'CS0009', name: 'ของเลิกใช้', qty: 50, unit: 'ตัว', active: false }
};
const ok = plan(stock, [{ item_id: 'CS0001', qty: 4 }, { item_id: 'CS0002', qty: 1 }, { item_id: 'CS0001', qty: 2 }]);
assert.strictEqual(ok.lines.length, 2, 'รายการซ้ำในตะกร้าต้องรวมเป็นแถวเดียว');
assert.strictEqual(ok.lines[0].qty, 6);
assert.strictEqual(ok.lines[0].balance, 4, 'ยอดคงเหลือหลังเบิก = 10 - 6');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 1 }, { item_id: 'CS0002', qty: 5 }]), /ของไม่พอเบิก: น็อต M8/, 'ของไม่พอต้องโยน error ทั้งชุด');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 6 }, { item_id: 'CS0001', qty: 6 }]), /ของไม่พอเบิก/, 'รวมยอดซ้ำแล้วเกินต้องถูกจับได้');
assert.throws(() => plan(stock, [{ item_id: 'CS0009', qty: 1 }]), /ไม่พบรายการ/, 'ของที่ปิดใช้งานต้องเบิกไม่ได้');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 0 }]), /จำนวนต้องมากกว่า 0/);
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: -2 }]), /จำนวนต้องมากกว่า 0/);

// ── เลขใบเบิกวิ่งใหม่ทุกวัน และไม่สับสนกับวันอื่น ─────────────────────────────────
const nextTicket = new Function(grab(backend, /^function nextConsumableTicketNo\(prefix, existingTickets\) \{[\s\S]*?\n\}/m, 'nextConsumableTicketNo') + '\nreturn nextConsumableTicketNo;')();
assert.strictEqual(nextTicket('CS-261029', []), 'CS-261029-001');
assert.strictEqual(nextTicket('CS-261029', ['CS-261028-009', 'CS-261029-002', '', 'CS-261029-011', 'CS-261029-003']), 'CS-261029-012');

// ── Backend: สิทธิ์ + เขียนใต้ Lock + routing ────────────────────────────────────
assert(/function withdrawConsumables\(payload\) \{\n  var user = requireWarehouseWriter\(/.test(backend), 'เบิกต้องผ่านด่าน requireWarehouseWriter (บัญชีดูอย่างเดียวเบิกไม่ได้)');
['restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(fn) {
  assert(new RegExp('function ' + fn + '\\(payload\\) \\{\\n  var user = requireConsumableKeeper\\(').test(backend), fn + ' ต้องเป็น Admin/Leader เท่านั้น');
});
['withdrawConsumables', 'restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(fn) {
  const body = grab(backend, new RegExp('^function ' + fn + '\\(payload\\) \\{[\\s\\S]*?\\n\\}', 'm'), fn);
  assert(body.includes('LockService.getScriptLock()') && body.includes('lock.releaseLock()'), fn + ' ต้องเขียนใต้ Lock กันเบิกชนกันแล้วยอดเพี้ยน');
});
assert(/function adjustConsumableStock[\s\S]{0,600}if \(!reason\) throw/.test(backend), 'นับปรับยอดต้องบังคับเหตุผล');
['getConsumables', 'withdrawConsumables', 'restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(action) {
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(e.parameter), e);"), action + ' ต้องเรียกผ่าน GET ได้');
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(body), e);"), action + ' ต้องเรียกผ่าน POST ได้');
});
assert(backend.includes("SPARE_APP_CONFIG.consumableSheetName = SPARE_APP_CONFIG.consumableSheetName || 'Consumables'"));
assert(backend.includes("SPARE_APP_CONFIG.consumableLogSheetName = SPARE_APP_CONFIG.consumableLogSheetName || 'ConsumableLog'"));
const kit = new Function(grab(backend, /^var CONSUMABLE_CATEGORIES = [^\n]*$/m, 'categories') + '\n' + grab(backend, /^var CONSUMABLE_STARTER_KIT = \[[\s\S]*?\n\];/m, 'starter kit') + '\nreturn { cats: CONSUMABLE_CATEGORIES, kit: CONSUMABLE_STARTER_KIT };')();
kit.kit.forEach(function(row) { assert(kit.cats.indexOf(row[1]) > -1, 'ชุดเริ่มต้นต้องใช้หมวดที่มีจริง: ' + row[1]); });

// ── หน้าเว็บ: ตัวช่วยคำนวณ ─────────────────────────────────────────────────────
const fe = new Function([
  grab(htmlLf, /^ {4}function csStockLevel\(item\) \{[\s\S]*?\n {4}\}/m, 'csStockLevel'),
  grab(htmlLf, /^ {4}function csForecastDaysLeft\(qty, issuedInWindow, windowDays\) \{[\s\S]*?\n {4}\}/m, 'csForecastDaysLeft'),
  grab(htmlLf, /^ {4}function csReorderQty\(item\) \{[\s\S]*?\n {4}\}/m, 'csReorderQty'),
  grab(htmlLf, /^ {4}function csGroupHistory\(logs\) \{[\s\S]*?\n {4}\}/m, 'csGroupHistory'),
  grab(htmlLf, /^ {4}function csMyFavorites\(logs, me, limit\) \{[\s\S]*?\n {4}\}/m, 'csMyFavorites')
].join('\n') + '\nreturn { csStockLevel, csForecastDaysLeft, csReorderQty, csGroupHistory, csMyFavorites };')();

assert.strictEqual(fe.csStockLevel({ qty: 0, min: 5, max: 20 }).level, 'out');
assert.strictEqual(fe.csStockLevel({ qty: 5, min: 5, max: 20 }).level, 'low', 'ถึง Min พอดีถือว่าใกล้หมด');
assert.strictEqual(fe.csStockLevel({ qty: 6, min: 5, max: 20 }).level, 'ok');
assert.strictEqual(fe.csStockLevel({ qty: 10, min: 5, max: 20 }).pct, 50);
assert.strictEqual(fe.csStockLevel({ qty: 3, min: 0, max: 0 }).level, 'ok', 'ไม่ตั้ง Min = ไม่เตือน');

assert.strictEqual(fe.csForecastDaysLeft(30, 30, 30), 30, 'ใช้วันละ 1 เหลือ 30 = อีก 30 วัน');
assert.strictEqual(fe.csForecastDaysLeft(10, 60, 30), 5);
assert.strictEqual(fe.csForecastDaysLeft(10, 0, 30), null, 'ไม่มีประวัติเบิกต้องไม่เดา');
assert.strictEqual(fe.csForecastDaysLeft(0, 10, 30), 0);

assert.strictEqual(fe.csReorderQty({ qty: 12, min: 20, max: 100 }), 88, 'สั่งให้เติมถึง Max');
assert.strictEqual(fe.csReorderQty({ qty: 3, min: 5, max: 0 }), 7, 'ไม่มี Max ใช้ Min×2');

const groups = fe.csGroupHistory([
  { txn_id: 'a', ticket_no: 'CS-1', type: 'ISSUE' }, { txn_id: 'b', ticket_no: 'CS-1', type: 'ISSUE' },
  { txn_id: 'c', ticket_no: '', type: 'RESTOCK' }, { txn_id: 'd', ticket_no: 'CS-2', type: 'ISSUE' }
]);
assert.deepStrictEqual(groups.map(g => g.key + ':' + g.logs.length), ['CS-1:2', 'c:1', 'CS-2:1'], 'ใบเบิกเดียวกันต้องรวมเป็นใบเดียว ส่วนเติม/ปรับยอดแสดงทีละรายการ');

assert.deepStrictEqual(fe.csMyFavorites([
  { type: 'ISSUE', item_id: 'A', requester: 'me' }, { type: 'ISSUE', item_id: 'B', recorded_by: 'me' }, { type: 'ISSUE', item_id: 'B', requester: 'me' },
  { type: 'ISSUE', item_id: 'C', requester: 'other', recorded_by: 'other' }, { type: 'RESTOCK', item_id: 'D', recorded_by: 'me' }
], 'me', 8), ['B', 'A'], 'เบิกบ่อยของฉันต้องนับเฉพาะของตัวเองและเฉพาะการเบิก');

// ── หน้าเว็บ: เมนู Sidebar + หน้า + สลับหน้า ──────────────────────────────────────
const sideAt = htmlLf.indexOf('id="csSideBtn"');
assert(sideAt > htmlLf.indexOf('id="lineTabs"') && sideAt < htmlLf.indexOf('id="lineMenuFooterAvatar"'), 'ปุ่มเบิกของสิ้นเปลืองต้องอยู่ใน Sidebar ใต้รายชื่อไลน์ เหนือข้อมูลผู้ใช้');
assert(htmlLf.includes('id="consumablePage"') && htmlLf.includes('id="csCheckoutModal"') && htmlLf.includes('id="csSuccessModal"'));
assert(htmlLf.includes("'consumables': { btn: 'csSideBtn'"), 'ต้องเปิดเป็นแท็บแบบ browser ได้');
assert((htmlLf.match(/^ {6}csHidePage\(\);$/gm) || []).length >= 13, 'ทุก switch page ต้องซ่อนหน้าเบิกของสิ้นเปลือง');
assert(/function switchToMiscExpensePage\(\)[\s\S]{0,800}csHidePage\(\);/.test(htmlLf), 'หน้าค่าใช้จ่ายสิ้นเปลืองต้องซ่อนหน้านี้ด้วย');
assert(/function switchToConsumablePage\(\)[\s\S]{0,700}mxHidePage\(\);[\s\S]{0,400}setTabStyles\('consumables'\)/.test(htmlLf));
assert(/function renderAuthUserBox\(\)[\s\S]{0,400}if \(!currentUser\) \{[\s\S]{0,120}csSyncSideEntry\(\);/.test(htmlLf), 'ออกจากระบบต้องล้างตะกร้าและซ่อนปุ่ม');
assert(/function csCanWithdraw\(\) \{ return canWriteWarehouse\(\); \}/.test(htmlLf), 'ฝั่งเว็บต้องสะท้อนด่าน requireWarehouseWriter');
assert(htmlLf.includes(".cs-cartbar.hidden") || htmlLf.includes('#consumablePage .hidden'), 'คลาส cs-* ที่ตั้ง display เองต้องยังซ่อนด้วย .hidden ได้');
