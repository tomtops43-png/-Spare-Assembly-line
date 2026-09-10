// ตัวกรองช่วงวันที่ของ Export Center — เลือกวันได้/เลือกช่วงได้ตามที่ผู้จัดการต้องการ
// จุดที่พลาดง่ายและเทสต์นี้กันไว้:
//  1) toISOString() เป็น UTC ตอนเช้าเวลาไทยจะได้วันย้อนหลังไป 1 วัน
//  2) Google Sheets แปลงคอลัมน์วันที่เป็น Date object เอง ต้องอ่านได้ทั้ง string และ Date
//  3) แถวที่ไม่มีวันที่ต้องไม่หายไปเงียบๆ จากไฟล์
const { html, backend, grabFn, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat(['var xpFilters = { from: "", to: "", line: "", groupBy: "month" };']),
  '{ xpDateStr: xpDateStr, xpDayKey: xpDayKey, xpMonthKey: xpMonthKey, xpWeekKey: xpWeekKey, xpBucketKey: xpBucketKey, xpInRange: xpInRange, xpLineMatches: xpLineMatches, setFilters: function(f) { xpFilters = f; } }'
);

// ── xpDateStr ต้องเป็นวันตามเวลาเครื่อง ไม่ใช่ UTC ───────────────────
// 2026-09-10 07:00 เวลาไทย = 2026-09-10 00:00Z ถ้าใช้ toISOString จะยังได้ 09-10
// แต่ 2026-09-10 06:00 เวลาไทย = 2026-09-09 23:00Z ซึ่ง toISOString จะให้ 09-09 (ผิด)
const early = new Date(2026, 8, 10, 6, 0, 0);
assert.strictEqual(mod.xpDateStr(early), '2026-09-10', 'ต้องได้วันตามเวลาเครื่อง ไม่ใช่ UTC');
assert.strictEqual(mod.xpDateStr(new Date(2026, 0, 1, 0, 30)), '2026-01-01');
assert.strictEqual(mod.xpDateStr(new Date(2026, 11, 31, 23, 59)), '2026-12-31');

// ── อ่านวัน/เดือน ได้จากทุกรูปแบบที่ระบบเก็บไว้จริง ───────────────────
assert.strictEqual(mod.xpDayKey('2026-09-10 14:26:47'), '2026-09-10', 'รูปแบบที่ processTransaction เขียนลง Log');
assert.strictEqual(mod.xpDayKey('2026-09-10'), '2026-09-10');
assert.strictEqual(mod.xpDayKey('10/9/2026'), '2026-09-10', 'รูปแบบ d/M/yyyy จากไฟล์นำเข้าเก่า');
assert.strictEqual(mod.xpDayKey(''), '', 'ค่าว่างต้องได้ค่าว่าง ไม่ใช่วันนี้');
assert.strictEqual(mod.xpMonthKey('2026-09-10 14:26:47'), '2026-09');
assert.strictEqual(mod.xpMonthKey('2026-09'), '2026-09');
assert.strictEqual(mod.xpMonthKey('10/9/2026'), '2026-09');
assert.strictEqual(mod.xpMonthKey(new Date(2026, 8, 1)), '2026-09', 'Date object จากชีตต้องอ่านได้');

// ── ช่วงวันที่ ───────────────────────────────────────────────────────
mod.setFilters({ from: '2026-09-01', to: '2026-09-10', line: '', groupBy: 'month' });
assert.strictEqual(mod.xpInRange('2026-09-05 08:00:00'), true);
assert.strictEqual(mod.xpInRange('2026-09-01 00:00:01'), true, 'วันเริ่มต้องรวมอยู่ในช่วง');
assert.strictEqual(mod.xpInRange('2026-09-10 23:59:59'), true, 'วันสิ้นสุดต้องรวมอยู่ในช่วง');
assert.strictEqual(mod.xpInRange('2026-08-31 23:00:00'), false);
assert.strictEqual(mod.xpInRange('2026-09-11 00:00:00'), false);
// แถวที่ไม่มีวันที่: ปล่อยผ่านให้เห็นในไฟล์ ดีกว่าตัดทิ้งเงียบๆ แล้วยอดรวมไม่ตรงชีต
assert.strictEqual(mod.xpInRange(''), true, 'แถวไม่มีวันที่ต้องไม่ถูกตัดทิ้ง');

// เลือก "ทั้งหมด" = ไม่กรองวันเลย
mod.setFilters({ from: '', to: '', line: '', groupBy: 'month' });
assert.strictEqual(mod.xpInRange('2019-01-01'), true, 'ไม่ตั้งช่วง = เอาทุกวัน');

// เลือกวันเดียว (from = to) ต้องได้เฉพาะวันนั้น
mod.setFilters({ from: '2026-09-10', to: '2026-09-10', line: '', groupBy: 'month' });
assert.strictEqual(mod.xpInRange('2026-09-10 09:00:00'), true);
assert.strictEqual(mod.xpInRange('2026-09-09 23:59:00'), false);

// ── กรองไลน์ ─────────────────────────────────────────────────────────
mod.setFilters({ from: '', to: '', line: 'Lug&Screw', groupBy: 'month' });
assert.strictEqual(mod.xpLineMatches('Lug&Screw'), true);
assert.strictEqual(mod.xpLineMatches('lug&screw'), true, 'ต้องไม่แคร์ตัวพิมพ์');
assert.strictEqual(mod.xpLineMatches('Lug & Screw'), true, 'ต้องไม่แคร์ช่องว่าง — Process ในชีต Log พิมพ์ไม่เหมือนกันทุกแถว');
assert.strictEqual(mod.xpLineMatches('H9'), false);
assert.strictEqual(mod.xpLineMatches(''), true, 'แถวที่ไม่ระบุไลน์ต้องไม่ถูกตัดทิ้ง');
mod.setFilters({ from: '', to: '', line: '', groupBy: 'month' });
assert.strictEqual(mod.xpLineMatches('อะไรก็ได้'), true, 'ไม่เลือกไลน์ = เอาทุกไลน์');

// ── รวมข้อมูลตามวัน / สัปดาห์ / เดือน ─────────────────────────────────
mod.setFilters({ from: '', to: '', line: '', groupBy: 'day' });
assert.strictEqual(mod.xpBucketKey('2026-09-10 14:00:00'), '2026-09-10');
mod.setFilters({ from: '', to: '', line: '', groupBy: 'month' });
assert.strictEqual(mod.xpBucketKey('2026-09-10 14:00:00'), '2026-09');
mod.setFilters({ from: '', to: '', line: '', groupBy: 'week' });
const w = mod.xpBucketKey('2026-09-10');
assert(/^\d{4}-W\d{2}$/.test(w), 'คีย์สัปดาห์ต้องเป็นรูปแบบ yyyy-Www แต่ได้ ' + w);
// วันในสัปดาห์เดียวกันต้องได้คีย์เดียวกัน (จ. 7 ก.ย. 2026 ถึง อา. 13 ก.ย. 2026)
assert.strictEqual(mod.xpBucketKey('2026-09-07'), mod.xpBucketKey('2026-09-13'), 'จันทร์กับอาทิตย์ของสัปดาห์เดียวกันต้องอยู่ช่องเดียวกัน');
assert.notStrictEqual(mod.xpBucketKey('2026-09-06'), mod.xpBucketKey('2026-09-07'), 'อาทิตย์ก่อนหน้าต้องคนละสัปดาห์');

// ── ปุ่มช่วงเวลาแบบกดเร็ว ต้องคำนวณครบทุกตัวเลือก ─────────────────────
const quickSrc = grabFn('xpApplyQuickRange');
['today', '7d', 'month', 'prevmonth', 'quarter', 'year'].forEach(function(k) {
  assert(quickSrc.indexOf("'" + k + "'") > -1, 'xpApplyQuickRange ต้องรองรับ ' + k);
});
// "เดือนก่อน" ต้องปิดท้ายที่วันสุดท้ายของเดือนก่อน ไม่ใช่วันนี้
assert(/prevmonth[\s\S]*now\.getMonth\(\), 0\)/.test(quickSrc), '"เดือนก่อน" ต้องตั้งวันสิ้นสุดเป็นวันสุดท้ายของเดือนนั้น');
// "7 วัน" ต้องนับรวมวันนี้ (6 วันย้อนหลัง + วันนี้) ไม่ใช่ 7 วันย้อนหลังกลายเป็น 8 วัน
assert(/6 \* 86400000/.test(quickSrc), '"7 วัน" ต้องย้อนหลัง 6 วันเพื่อให้นับรวมวันนี้ครบ 7 วัน');

// ── กรอกช่วงกลับหัวต้องสลับให้เอง ไม่ใช่คืนไฟล์เปล่า ────────────────
const readSrc = grabFn('xpReadFilterInputs');
assert(/xpFilters\.from > xpFilters\.to/.test(readSrc), 'ต้องตรวจว่าผู้ใช้กรอกวันกลับหัว');
assert(/สลับวันเริ่ม/.test(readSrc), 'สลับให้แล้วต้องบอกผู้ใช้ด้วย');
// เปลี่ยนตัวกรองแล้วต้องล้าง cache ไม่งั้นจะได้ไฟล์ของช่วงเดิม
const afterSrc = grabFn('xpAfterFilterChange');
assert(/xpClearCache\(\)/.test(afterSrc), 'เปลี่ยนตัวกรองต้องล้าง cache ข้อมูลที่ดึงไว้');

// ── ฝั่งเซิร์ฟเวอร์ต้องกรองช่วงวันที่ด้วย ไม่ใช่พึ่งฝั่งเว็บอย่างเดียว ──
assert(/function exportDateKey\(/.test(backend), 'Backend ต้องมี exportDateKey');
assert(/function exportInDateRange\(/.test(backend), 'Backend ต้องมี exportInDateRange');
assert(/function getLogRows\(options\)/.test(backend), 'getLogRows ต้องรับตัวเลือกช่วงวันที่ (ชีต Log โตเรื่อยๆ ดึงทั้งก้อนจะชน timeout)');
assert(/if \(action === 'logs'\) return respond\(getLogRows\(\{ from: e\.parameter\.from/.test(backend), 'doGet action=logs ต้องส่งช่วงวันที่ต่อไปให้ getLogRows');
// ฝั่งเว็บต้องกรองซ้ำอีกชั้น เผื่อ Apps Script ยัง deploy ไม่ทันแล้วส่งทั้งชีตกลับมา
const fetchLogsSrc = grabFn('xpFetchLogs');
assert(/xpInRange\(l\.timestamp\)/.test(fetchLogsSrc), 'xpFetchLogs ต้องกรองช่วงวันซ้ำฝั่งเว็บ เผื่อ backend รุ่นเก่ายังไม่รู้จัก from/to');

console.log('export-filter-range: OK');
