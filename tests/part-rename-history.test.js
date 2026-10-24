// แก้ชื่ออะไหล่ใน Master แล้วประวัติการเบิกเดิมต้องไม่หาย
// เคสจริง: ผู้ใช้อยากเปลี่ยน "ไฟสี่เหลี่ยม" (รุ่น CA-DQW12M, H9) เป็น "Camera square flash"
// แต่ Log เก็บชื่อ ณ วันที่เบิก → ถ้าจับคู่ด้วย ชื่อ+รุ่น อย่างเดียว มูลค่าเบิกย้อนหลังกลายเป็น 0
// (หลุดจากอันดับใน PowerPoint / กราฟรายเดือน / Dashboard) เทสต์นี้กันไว้
const { html, grabFn, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", category: "", txnType: "", groupBy: "month", lang: "th" };',
    grabFn('fccN'), grabFn('fccPrice'), grabFn('fccLineOf'), grabFn('fccPartKey'), grabFn('fccModelKey'),
    grabFn('fccCanonicalLogs'), grabFn('fccCrunchLogs'),
    grabFn('xpDeckMoney'), grabFn('xpDeckMonthLabel'), grabFn('xpDeckMonthlyIssue')
  ]),
  '{ canon: fccCanonicalLogs, crunchLogs: fccCrunchLogs, partKey: fccPartKey, priceMap: xpPriceMap, lookup: xpLookupPrice, monthly: xpDeckMonthlyIssue }'
);

const master = [
  { line: 'H9', name: 'Camera square flash', model: 'CA-DQW12M', unit_price: 35000 },   // เปลี่ยนชื่อแล้ว
  { line: 'H9', name: 'Bolt M8', model: 'M8', unit_price: 2 },
  { line: 'Lug&Screw', name: 'Bolt M8 long', model: 'M8', unit_price: 3 },              // รุ่นสั้น/ซ้ำข้ามไลน์
  { line: 'H9', name: 'Sensor A', model: 'SN-100', unit_price: 500 },
  { line: 'H9', name: 'Sensor B', model: 'SN-100', unit_price: 700 },                    // รุ่นเดียวสองชื่อ
  { line: 'H9', name: 'Gasket', model: '-', unit_price: 10 }
];
const logs = [
  { timestamp: '2026-10-03 08:00:00', type: 'Output', process: 'H9', partName: 'ไฟสี่เหลี่ยม', model: 'CA-DQW12M', qty: 1 },   // ชื่อเก่า
  { timestamp: '2026-10-05 08:00:00', type: 'Output', process: 'H9', partName: 'Camera square flash', model: 'CA-DQW12M', qty: 1 },
  { timestamp: '2026-10-06 08:00:00', type: 'Output', process: 'H9', partName: 'Sensor old', model: 'SN-100', qty: 1 },        // รุ่นกำกวม
  { timestamp: '2026-10-07 08:00:00', type: 'Output', process: 'Lug&Screw', partName: 'ไฟสี่เหลี่ยม', model: 'CA-DQW12M', qty: 1 }, // คนละไลน์
  { timestamp: '2026-10-08 08:00:00', type: 'Output', process: 'H9', partName: 'Old gasket', model: '-', qty: 1 }               // รุ่นใช้ไม่ได้
];

// ── จับคู่ Log ชื่อเก่า → ชื่อปัจจุบัน ─────────────────────────────────
const c = mod.canon(logs, master);
assert.strictEqual(c[0].partName, 'Camera square flash', 'ชื่อเก่า + รุ่นไม่ซ้ำในไลน์ → ใช้ชื่อปัจจุบัน');
assert.strictEqual(c[0].partNameAtIssue, 'ไฟสี่เหลี่ยม', 'เก็บชื่อ ณ วันที่เบิกไว้ด้วย');
assert.strictEqual(c[1], logs[1], 'ตรงอยู่แล้วไม่ต้องแตะ');
assert.strictEqual(c[2].partName, 'Sensor old', 'รุ่นเดียวมีหลายชื่อ → เดาไม่ได้ ห้ามจับผิดตัว');
assert.strictEqual(c[3].partName, 'ไฟสี่เหลี่ยม', 'คนละไลน์ ไม่จับคู่ข้ามไลน์');
assert.strictEqual(c[4].partName, 'Old gasket', 'รุ่น "-" ใช้ระบุตัวไม่ได้');
assert.strictEqual(logs[0].partName, 'ไฟสี่เหลี่ยม', 'ห้ามแก้ข้อมูลต้นฉบับ (คืนสำเนา)');

// ── Dashboard: ยอดเบิก 30 วันต้องรวมเป็นตัวเดียว ─────────────────────
const now = new Date();
function d(daysAgo) { const x = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysAgo, 9); return x.toISOString().slice(0, 10) + ' 09:00:00'; }
const recent = [
  { timestamp: d(2), type: 'Output', process: 'H9', partName: 'ไฟสี่เหลี่ยม', model: 'CA-DQW12M', qty: 2 },
  { timestamp: d(1), type: 'Output', process: 'H9', partName: 'Camera square flash', model: 'CA-DQW12M', qty: 1 }
];
const L = mod.crunchLogs(mod.canon(recent, master), 'all');
const u = L.usage30[mod.partKey('H9', 'Camera square flash', 'CA-DQW12M')];
assert(u && u.qty === 3 && u.count === 2, 'ยอดเบิก 30 วันของชื่อใหม่ต้องรวมประวัติชื่อเก่าด้วย');

// ── ราคา (Export/Excel/PowerPoint) ตกไปใช้รุ่นได้เมื่อรุ่นไม่กำกวม ─────
const pm = mod.priceMap(master);
assert.strictEqual(mod.lookup(pm, 'ไฟสี่เหลี่ยม', 'CA-DQW12M', '', ''), 35000, 'ชื่อเก่ายังได้ราคาจากรุ่น');
assert.strictEqual(mod.lookup(pm, 'Sensor old', 'SN-100', '', ''), 0, 'รุ่นกำกวม → ไม่เดาราคา');
assert.strictEqual(mod.lookup(pm, 'Old bolt', 'M8', '', ''), 0, 'รุ่นเดียวหลายชื่อ (ข้ามไลน์) → ไม่เดาราคา');
assert.strictEqual(mod.lookup(pm, 'Old gasket', '-', '', ''), 0, 'รุ่น "-" → ไม่เดาราคา');
assert.strictEqual(mod.lookup(pm, 'Camera square flash', 'CA-DQW12M', '', ''), 35000);

// ── PowerPoint: อันดับมูลค่าเบิกต้องเป็นแถวเดียวในชื่อใหม่ ──────────────
const M = mod.monthly(logs.slice(0, 2), master, [], new Date(2026, 9, 24));
assert.strictEqual(M.rankAll.length, 1, 'ชื่อเก่า + ชื่อใหม่ ต้องรวมเป็นแถวเดียว ไม่แตกเป็นสองอันดับ');
assert.strictEqual(M.rankAll[0].name, 'Camera square flash');
assert.strictEqual(M.rankAll[0].value, 70000, 'มูลค่ารวมทั้งก่อนและหลังเปลี่ยนชื่อ');
assert.strictEqual(M.rankAll[0].count, 2);

// ── จุดที่ต้องผ่านตัวจับคู่ชื่อ ─────────────────────────────────────────
assert(/fccCrunchLogs\(fccCanonicalLogs\(A\.logs, all\), selLine\)/.test(html), 'หน้า Dashboard ต้องจับคู่ชื่อก่อนคำนวณ');
assert(/return fccCanonicalLogs\(logs, partsList\)\.map/.test(grabFn('fccExpenseValidRows')), 'ยอดเบิกตีมูลค่าบน Dashboard ต้องจับคู่ชื่อด้วย');
assert(/res\[1\] = fccCanonicalLogs\(res\[1\] \|\| \[\], res\[0\] \|\| \[\]\)/.test(grabFn('xpDashboardData')), 'Export/PowerPoint ต้องจับคู่ชื่อก่อนคำนวณ');

console.log('part-rename-history: OK');
