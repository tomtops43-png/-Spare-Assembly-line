// การ์ดหน้า Stock ขึ้นป้าย "🚚 กำลังสั่ง" เมื่ออะไหล่ตัวนั้นมี PR ค้างรับในระบบติดตาม PR
const { grabFn, grabVar, assert } = require('./_export-extract');

const helpers = ['normPhText', 'normPhModel', 'isMeaningfulPhModel', 'ptFmtNum', 'ptOnOrderLinesForItem', 'ptOnOrderBadgeHtml'].map(function(n) { return grabFn(n); }).join('\n');
function build(open) {
  return new Function('ptData', 'escHtml', grabVar('PT_STATUS') + ';\n' + helpers + '\nreturn { lines: ptOnOrderLinesForItem, badge: ptOnOrderBadgeHtml };')(
    { open: open }, function(v) { return String(v); });
}

const open = [
  { pr_id: 'PR-2610-001', line_no: 1, part_name: 'Relay', model: 'RXM4AB2BD', qty_outstanding: 4, unit: 'PCS', pr_status: 'ORDERED', days_waiting: 3 },
  { pr_id: 'PR-2610-002', line_no: 2, part_name: 'Relay', model: 'RXM-4AB2BD', qty_outstanding: 2, unit: 'PCS', pr_status: 'APPROVED', days_waiting: 1 },
  { pr_id: 'PR-2610-003', line_no: 1, part_name: 'Spring  Balancer', model: '-', qty_outstanding: 1, unit: 'PCS', pr_status: 'PARTIAL', days_waiting: 20 }
];
const api = build(open);

// จับคู่ด้วยรุ่น (ไม่สนขีด/ช่องว่าง) — รุ่นต่างกันแม้ชื่อเหมือนกันต้องไม่ขึ้น
assert.strictEqual(api.lines({ name: 'Relay', model: 'RXM4AB2BD' }).length, 2);
assert.strictEqual(api.lines({ name: 'Relay', model: 'RXM2AB2BD' }).length, 0, 'Relay คนละรุ่นต้องไม่ขึ้นกำลังสั่ง');
// ไม่มีรุ่น → เทียบชื่อแบบไม่สนช่องว่าง
assert.strictEqual(api.lines({ name: 'Spring Balancer', model: '-' }).length, 1);

const badge = api.badge({ name: 'Relay', model: 'RXM4AB2BD' }, false);
assert(badge.includes('กำลังสั่ง 6 PCS'), 'รวมยอดค้างรับทุก PR');
assert(badge.includes('PR-2610-001 +1') && badge.includes('สั่งซื้อแล้ว'), 'บอกเลข PR และสถานะ');
assert(badge.includes('data-onorder-pr="PR-2610-001"'), 'กดป้ายแล้วเปิดรายละเอียด PR');
assert.strictEqual(api.badge({ name: 'Zebra Printhead', model: 'P1083320-016' }, false), '', 'ไม่มี PR ค้าง = ไม่มีป้าย');
assert.strictEqual(build(null).badge({ name: 'Relay', model: 'RXM4AB2BD' }, true), '', 'ยังไม่โหลดข้อมูล = ไม่มีป้าย');

// ป้ายแสดงทั้งการ์ดเดสก์ท็อปและมือถือ + โหลดข้อมูลตอน render หน้า Stock
assert(grabFn('renderTable').includes('ptOnOrderBadgeHtml(item, false)'));
assert(grabFn('renderMobileCards').includes('ptOnOrderBadgeHtml(item, true)'));
const rf = grabFn('renderFilteredResults');
assert(rf.includes('ptEnsureOnOrderLoaded()') && rf.includes('bindOnOrderBadges(desktopCardGrid)') && rf.includes('bindOnOrderBadges(mobileCardList)'));
const ensure = grabFn('ptEnsureOnOrderLoaded');
assert(ensure.includes("action: 'getOpenPrLines'") && ensure.includes('60000'), 'ดึงรายการค้างรับ และแคชไว้ 1 นาที');

console.log('stock-card-on-order-badge: OK');
