// Dashboard Snapshot — เอาหน้า Factory Command Center ไปอยู่ในไฟล์ที่ส่งออก
// สิ่งที่เทสต์นี้กันไว้เป็นหลัก: **ตัวเลขในไฟล์ต้องตรงกับที่เห็นบนหน้า Dashboard เป๊ะ**
// วิธีกัน: บังคับให้ใช้ฟังก์ชันคำนวณของ Dashboard เอง (fccCrunchParts / fccCrunchLogs /
// fccCrunchRequests / fccHealth) แล้วเทียบผลลัพธ์ที่ประกอบขึ้นกับผลจากฟังก์ชันเดิมตรงๆ
// ถ้าวันหลังมีคนไปเขียนสูตรใหม่ในฝั่ง export ตัวเลขจะเริ่มไม่ตรงกัน เทสต์นี้จะจับได้
const { html, grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

// ── ต้องเรียกใช้ตัวคำนวณของ Dashboard ไม่ใช่เขียนสูตรใหม่ ────────────
const dashDataSrc = grabFn('xpDashboardData');
['fccCrunchParts', 'fccCrunchLogs', 'fccCrunchRequests', 'fccHealth'].forEach(function(fn) {
  assert(dashDataSrc.indexOf(fn) > -1,
    'xpDashboardData ต้องเรียก ' + fn + ' ของหน้า Dashboard — ห้ามคำนวณเองเพราะเลขจะไม่ตรงกับหน้าจอ');
});
assert(/typeof fccCrunchParts !== 'function'/.test(dashDataSrc),
  'ต้องเช็คว่าตัวคำนวณของ Dashboard มีอยู่จริง แล้วแจ้ง error ที่อ่านรู้เรื่อง ไม่ใช่พังเงียบๆ');

// KPI ของ Dashboard เป็นสถานะ ณ ปัจจุบัน ต้องไม่ถูกตัดด้วยช่วงวันที่ที่เลือกในหน้า Export
const snapLogsSrc = grabFn('xpFetchLogsSnapshot');
assert(!/xpFilters\.from/.test(snapLogsSrc) && !/xpInRange/.test(snapLogsSrc),
  'xpFetchLogsSnapshot ต้องไม่กรองช่วงวันที่ — ไม่งั้นเลือก "เดือนก่อน" แล้ว "เบิกวันนี้" จะขึ้น 0 ทั้งที่วันนี้มีเบิก');
assert(/xpLineMatches/.test(snapLogsSrc), 'ตัวกรองไลน์ยังต้องมีผล (Dashboard ก็มีตัวเลือกไลน์ของตัวเอง)');
const snapReqSrc = grabFn('xpFetchOrderRequestsSnapshot');
assert(!/xpInRange/.test(snapReqSrc), 'คำขอซื้อของ Dashboard เป็นยอดค้างปัจจุบัน ต้องไม่กรองด้วยวันที่');
// ต้อง cache แยกจากชุดที่กรองวันที่ ไม่งั้นสองชุดทับกัน
assert(/xpCached\('logsSnapshot'/.test(snapLogsSrc), 'ต้อง cache แยกจาก logs ที่กรองวันที่แล้ว');
assert(/xpCached\('orderRequestsSnapshot'/.test(snapReqSrc), 'ต้อง cache แยกจาก orderRequests ที่กรองวันที่แล้ว');
// ctx ถูกกรองไลน์มาแล้ว ต้องไม่ให้ fccCrunch* กรองซ้ำอีกชั้น (จะได้ศูนย์)
assert(/fccCrunchLogs\(res\[1\] \|\| \[\], 'all'\)/.test(dashDataSrc), "ต้องส่ง 'all' ให้ fccCrunchLogs เพราะกรองไลน์ไปแล้วรอบหนึ่ง");
assert(/fccCrunchRequests\(res\[2\] \|\| \[\], 'all'\)/.test(dashDataSrc), "ต้องส่ง 'all' ให้ fccCrunchRequests");

// ── ประกอบโมดูลจริง: ใช้ fccCrunch* ตัวจริงจาก index.html ────────────
const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", category: "", txnType: "", groupBy: "month", lang: "th" };',
    'var currentUser = { username: "manager01", role: "admin" };',
    'var fccUid = 0;',
    grabFn('fccN'), grabFn('fccPrice'), grabFn('fccInt'), grabFn('fccMoney'),
    grabFn('fccLineOf'), grabFn('fccPartKey'),
    grabFn('fccCrunchParts'), grabFn('fccCrunchLogs'), grabFn('fccCrunchRequests'), grabFn('fccHealth'),
    grabVar('XP_CHART_COLORS'), grabVar('XP_CHART_SPECS'),
    grabFn('xpColIdx'), grabFn('xpChartSeries'), grabFn('xpChartConfig'),
    grabFn('xpSeriesFromItems'),
    grabFn('xpDashKpiRows'), grabFn('xpDashKpiTable'), grabFn('xpDashHealthTable'),
    grabFn('xpDashMovementTable'), grabFn('xpDashLineTable'), grabFn('xpDashCategoryTable'),
    grabFn('xpDashRequestTable'), grabFn('xpDashChartSeries'), grabFn('xpDashSheets')
  ]),
  '{ fccCrunchParts: fccCrunchParts, fccCrunchLogs: fccCrunchLogs, fccCrunchRequests: fccCrunchRequests, fccHealth: fccHealth, xpDashKpiRows: xpDashKpiRows, xpDashKpiTable: xpDashKpiTable, xpDashHealthTable: xpDashHealthTable, xpDashMovementTable: xpDashMovementTable, xpDashLineTable: xpDashLineTable, xpDashCategoryTable: xpDashCategoryTable, xpDashRequestTable: xpDashRequestTable, xpDashChartSeries: xpDashChartSeries, xpDashSheets: xpDashSheets, setLang: function(l) { xpFilters.lang = l; }, setCustomer: function(v) { xpFilters.customerMode = v; } }'
);

// ── ข้อมูลตัวอย่าง ────────────────────────────────────────────────────
const today = new Date();
function stamp(daysAgo, hh) {
  const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysAgo, hh || 9, 0, 0);
  const p = function(n) { return ('0' + n).slice(-2); };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':00:00';
}
const parts = [
  { name: 'เบรกเกอร์', model: 'BK-10', line: 'Lug&Screw', category: 'ไฟฟ้า', stock: 0, min: 10, unit_price: 100 },
  { name: 'น็อต M8', model: 'M8', line: 'Lug&Screw', category: 'ฮาร์ดแวร์', stock: 200, min: 50, unit_price: 2 },
  { name: 'สายพาน', model: 'B-77', line: 'H9', category: 'กลไก', stock: 1, min: 5, unit_price: 500 },
  { name: 'ฟิวส์', model: 'FZ-1', line: 'H9', category: 'ไฟฟ้า', stock: 20, min: 0, unit_price: '' }
];
const logs = [
  { timestamp: stamp(0), type: 'Output', process: 'Lug&Screw', partName: 'เบรกเกอร์', model: 'BK-10', qty: 6, by: 'a' },
  { timestamp: stamp(0), type: 'Input', process: 'H9', partName: 'สายพาน', model: 'B-77', qty: 4, by: 'b' },
  { timestamp: stamp(3), type: 'Output', process: 'H9', partName: 'สายพาน', model: 'B-77', qty: 3, by: 'b' },
  { timestamp: stamp(10), type: 'Output', process: 'Lug&Screw', partName: 'น็อต M8', model: 'M8', qty: 50, by: 'a' }
];
const reqs = [
  { request_id: 'RQ1', line: 'H9', status: 'Pending', requested_date: stamp(20).slice(0, 10) },
  { request_id: 'RQ2', line: 'H9', status: 'Pending', requested_date: stamp(2).slice(0, 10) },
  { request_id: 'RQ3', line: 'Lug&Screw', status: 'Approved', requested_date: stamp(5).slice(0, 10) }
];

const P = mod.fccCrunchParts(parts);
const L = mod.fccCrunchLogs(logs, 'all');
const R = mod.fccCrunchRequests(reqs, 'all');
const S = { P: P, L: L, R: R, H: mod.fccHealth(P, R) };

// ── การ์ด KPI ต้องตรงกับผลของ fccCrunch* ทุกช่อง ─────────────────────
const kpi = {};
mod.xpDashKpiRows(S).forEach(function(r) { kpi[r[0]] = r[1]; });
assert.strictEqual(kpi['อะไหล่ทั้งหมด'], P.total, 'จำนวนอะไหล่ต้องตรงกับ fccCrunchParts');
assert.strictEqual(kpi['มูลค่าคงคลัง'], Math.round(P.value), 'มูลค่าคงคลังต้องตรงกับ fccCrunchParts');
assert.strictEqual(kpi['หมดสต็อก'], P.out, 'หมดสต็อกต้องตรงกับ fccCrunchParts');
assert.strictEqual(kpi['ต่ำกว่า Min'], P.low, 'ต่ำกว่า Min ต้องตรงกับ fccCrunchParts');
assert.strictEqual(kpi['Need PO'], P.need, 'Need PO ต้องตรงกับ fccCrunchParts');
assert.strictEqual(kpi['PR รอดำเนินการ'], R.counts['Pending'], 'PR รอดำเนินการต้องตรงกับ fccCrunchRequests');
assert.strictEqual(kpi['เบิกวันนี้'], L.outToday, 'เบิกวันนี้ต้องตรงกับ fccCrunchLogs');
assert.strictEqual(kpi['รับเข้าวันนี้'], L.inToday, 'รับเข้าวันนี้ต้องตรงกับ fccCrunchLogs');
assert.strictEqual(kpi['Factory Health Score'], S.H.score, 'คะแนนสุขภาพต้องตรงกับ fccHealth');

// ค่าที่คาดไว้จากข้อมูลตัวอย่าง (ยืนยันว่าสูตรของ Dashboard ทำงานจริง ไม่ใช่ 0 ทั้งแถว)
assert.strictEqual(P.total, 4);
assert.strictEqual(P.out, 1, 'เบรกเกอร์ stock 0 = หมดสต็อก');
assert.strictEqual(P.low, 1, 'สายพาน stock 1 < min 5 = ต่ำกว่า Min');
assert.strictEqual(P.need, 2);
assert.strictEqual(L.outToday, 6, 'วันนี้เบิกเบรกเกอร์ 6 ชิ้น');
assert.strictEqual(L.inToday, 4, 'วันนี้รับสายพาน 4 ชิ้น');
assert.strictEqual(R.counts['Pending'], 2);
assert.strictEqual(R.over7, 1, 'RQ1 ค้าง 20 วัน = เกิน 7 วัน 1 คำขอ');
assert(P.value > 0, 'ต้องมีมูลค่าคงคลัง (ฟิวส์ที่ไม่มีราคาต้องไม่ถูกนับ)');
assert.strictEqual(P.pricedCount, 3, 'ฟิวส์ไม่มีราคา ต้องไม่นับเป็นรายการที่มีราคา');

// ── ที่มาของคะแนนสุขภาพต้องอธิบายได้ ─────────────────────────────────
const healthRows = {};
mod.xpDashHealthTable(S).rows.forEach(function(r) { healthRows[r[0]] = r[1]; });
assert.strictEqual(healthRows['คะแนนสุขภาพคลัง'], S.H.score);
assert.strictEqual(healthRows['ระดับ'], S.H.tone.g);
assert.strictEqual(healthRows['หักจากของหมดสต็อก'], -S.H.dOut, 'ต้องแสดงเป็นค่าลบให้เห็นว่าหักไป');
assert.strictEqual(healthRows['หักจากของต่ำกว่า Min'], -S.H.dLow);
assert.strictEqual(healthRows['หักจาก PR ที่ค้างอยู่'], -S.H.dPr);
// คะแนนต้องเท่ากับ 100 หักด้วยผลรวมที่แสดงไว้ (ปัดเศษได้ไม่เกิน 1 หน่วย)
const sumDeduct = S.H.dOut + S.H.dLow + S.H.dPr;
assert(Math.abs((100 - sumDeduct) - S.H.score) <= 1.5,
  'คะแนนต้องอธิบายได้จากยอดหักที่แสดง (100 - ' + sumDeduct + ' ≈ ' + S.H.score + ')');

// ── ตารางเคลื่อนไหว 14 วัน ต้องตรงกับ L.days ──────────────────────────
const mv = mod.xpDashMovementTable(S);
assert.strictEqual(mv.rows.length, 14, 'ต้องมี 14 แถว (14 วัน)');
assert.strictEqual(mv.rows.length, L.days.length);
assert.strictEqual(mv.rows[13][1], L.days[13].out, 'วันสุดท้าย (วันนี้) ต้องตรงกับ fccCrunchLogs');
assert.strictEqual(mv.rows[13][2], L.days[13].in);
const mvOutSum = mv.rows.reduce(function(a, r) { return a + r[1]; }, 0);
assert.strictEqual(mvOutSum, L.out14, 'ยอดรวมเบิกใน 14 วันต้องตรงกับ L.out14');

// ── ตารางตามไลน์ ต้องตรงกับ P.lines และรวมกันได้เท่ายอดรวม ────────────
const lineTable = mod.xpDashLineTable(S);
const liLine = lineTable.headers.indexOf('ไลน์');
const liTotal = lineTable.headers.indexOf('จำนวน SKU');
const liOut = lineTable.headers.indexOf('หมดสต็อก');
assert(lineTable.rows.length >= 2, 'ต้องมีอย่างน้อย 2 ไลน์');
const sumSku = lineTable.rows.reduce(function(a, r) { return a + r[liTotal]; }, 0);
assert.strictEqual(sumSku, P.total, 'ผลรวม SKU ทุกไลน์ต้องเท่ากับยอดรวม');
const sumOut = lineTable.rows.reduce(function(a, r) { return a + r[liOut]; }, 0);
assert.strictEqual(sumOut, P.out, 'ผลรวมของหมดทุกไลน์ต้องเท่ากับยอดรวม');
const h9Row = lineTable.rows.filter(function(r) { return r[liLine] === 'H9'; })[0];
assert(h9Row, 'ต้องมีแถวของไลน์ H9');
assert.strictEqual(h9Row[liTotal], P.lines['H9'].total);

// ── หมวดที่มีของเสี่ยง — เฉพาะหมวดที่มีของเสี่ยงจริง เรียงมากไปน้อย ───
const catTable = mod.xpDashCategoryTable(S);
assert(catTable.rows.length > 0, 'ข้อมูลตัวอย่างมีของเสี่ยง ต้องมีแถว');
catTable.rows.forEach(function(r) { assert(r[1] > 0, 'ห้ามมีหมวดที่เสี่ยง 0 รายการติดมา'); });
for (let i = 1; i < catTable.rows.length; i += 1) {
  assert(catTable.rows[i - 1][1] >= catTable.rows[i][1], 'ต้องเรียงจากเสี่ยงมากไปน้อย');
}
const sumRisk = catTable.rows.reduce(function(a, r) { return a + r[1]; }, 0);
assert.strictEqual(sumRisk, P.out + P.low, 'ผลรวมของเสี่ยงทุกหมวดต้องเท่ากับ หมด + ต่ำกว่า Min');

// ── สถานะคำขอซื้อ ─────────────────────────────────────────────────────
const reqTable = mod.xpDashRequestTable(S);
const reqMap = {};
reqTable.rows.forEach(function(r) { reqMap[r[0]] = r[1]; });
assert.strictEqual(reqMap['Pending'], 2);
assert.strictEqual(reqMap['Approved'], 1);

// ── กราฟของ Dashboard ────────────────────────────────────────────────
const dashCharts = mod.xpDashChartSeries(S);
assert(dashCharts.length >= 5, 'Dashboard ควรมีกราฟอย่างน้อย 5 ตัว แต่ได้ ' + dashCharts.length);
dashCharts.forEach(function(c) {
  assert(c.title && !/undefined/.test(c.title), 'กราฟ Dashboard ต้องมีหัวข้อ');
  assert(c.series && c.series.items.length, 'กราฟ "' + c.title + '" ไม่มีข้อมูล');
  c.series.items.forEach(function(it) {
    assert(typeof it.value === 'number' && isFinite(it.value), 'กราฟ "' + c.title + '" มีค่าที่ไม่ใช่ตัวเลข');
  });
});
const dashTitles = dashCharts.map(function(c) { return c.title; });
assert(dashTitles.indexOf('เบิกออก 14 วันล่าสุด (ชิ้น)') > -1, 'ต้องมีกราฟเบิก 14 วัน');
assert(dashTitles.indexOf('มูลค่าคงคลังตามไลน์') > -1, 'ต้องมีกราฟมูลค่าตามไลน์');
assert(dashTitles.indexOf('สถานะคำขอซื้อ') > -1, 'ต้องมีกราฟสถานะคำขอ');
// กราฟเบิก 14 วันต้องมี 14 แท่งเสมอ (รวมวันที่ไม่มีรายการ เพื่อให้เห็นช่องว่างจริง)
const outChart = dashCharts.filter(function(c) { return c.title === 'เบิกออก 14 วันล่าสุด (ชิ้น)'; })[0];
assert.strictEqual(outChart.series.items.length, 14, 'กราฟ 14 วันต้องมี 14 แท่ง แม้บางวันเป็น 0');
// สถานะคำขอเป็นกราฟวงกลม
const reqChart = dashCharts.filter(function(c) { return c.title === 'สถานะคำขอซื้อ'; })[0];
assert.strictEqual(reqChart.series.spec.type, 'doughnut', 'สถานะคำขอควรเป็นวงกลม (สัดส่วนของทั้งหมด)');

// ── โหมดลูกค้า: ต้องไม่มีมูลค่าเลย ทั้งการ์ด ตาราง และกราฟ ────────────
mod.setCustomer(true);
const custKpi = mod.xpDashKpiRows(S).map(function(r) { return r[0]; });
assert(custKpi.indexOf('มูลค่าคงคลัง') === -1, 'โหมดลูกค้าต้องไม่มีการ์ดมูลค่าคงคลัง');
assert(custKpi.indexOf('หมดสต็อก') > -1, 'การ์ดที่ไม่ใช่เรื่องเงินยังต้องอยู่');
const custLine = mod.xpDashLineTable(S);
assert.strictEqual(custLine.headers.indexOf('มูลค่าคงคลัง'), -1, 'ตารางตามไลน์ต้องไม่มีคอลัมน์มูลค่า');
const custCharts = mod.xpDashChartSeries(S).map(function(c) { return c.title; });
assert(custCharts.indexOf('มูลค่าคงคลังตามไลน์') === -1, 'โหมดลูกค้าต้องไม่มีกราฟมูลค่า');
assert(custCharts.indexOf('รายการที่ต้องสั่งซื้อตามไลน์') > -1, 'กราฟที่ไม่ใช่เรื่องเงินยังต้องอยู่');
mod.setCustomer(false);

// ── ภาษาอังกฤษ: ทุกป้ายในตาราง Dashboard ต้องถูกแปล ──────────────────
const THAI = /[฀-๿]/;
mod.setLang('en');
[mod.xpDashKpiTable(S), mod.xpDashHealthTable(S), mod.xpDashMovementTable(S), mod.xpDashLineTable(S), mod.xpDashCategoryTable(S), mod.xpDashRequestTable(S)].forEach(function(t, i) {
  (t.headers || []).forEach(function(h) {
    assert(!THAI.test(String(h)), 'ตาราง Dashboard #' + i + ' ยังมีหัวคอลัมน์ภาษาไทย: ' + h);
  });
});
mod.xpDashKpiTable(S).rows.forEach(function(r) {
  assert(!THAI.test(String(r[0])), 'ชื่อการ์ด KPI ยังไม่ถูกแปล: ' + r[0]);
  assert(!THAI.test(String(r[2])), 'หน่วยของ KPI ยังไม่ถูกแปล: ' + r[2]);
});
mod.xpDashHealthTable(S).rows.forEach(function(r) {
  assert(!THAI.test(String(r[0])), 'หัวข้อคะแนนสุขภาพยังไม่ถูกแปล: ' + r[0]);
  assert(!THAI.test(String(r[2])), 'คำอธิบายคะแนนสุขภาพยังไม่ถูกแปล: ' + r[2]);
});
mod.xpDashChartSeries(S).forEach(function(c) {
  assert(!THAI.test(c.title), 'หัวข้อกราฟ Dashboard ยังไม่ถูกแปล: ' + c.title);
});
mod.xpDashSheets(S).forEach(function(sh) {
  assert(!THAI.test(String(sh.name)), 'ชื่อชีต Dashboard ยังไม่ถูกแปล: ' + sh.name);
});
mod.setLang('th');

// ── ชีตของ Dashboard ที่ไปอยู่ใน Executive Workbook ──────────────────
const sheets = mod.xpDashSheets(S);
assert.strictEqual(sheets.length, 6, 'ต้องมีชีต Dashboard 6 ใบ');
sheets.forEach(function(sh) {
  assert(sh.name, 'ชีตต้องมีชื่อ');
  assert(sh.table && Array.isArray(sh.table.headers), 'ชีตต้องมีตาราง');
});
const names = sheets.map(function(s) { return s.name; });
assert(names.indexOf('05 Dashboard KPI') > -1);
assert(names.indexOf('06 คะแนนสุขภาพคลัง') > -1);
assert(names.indexOf('10 สถานะคำขอซื้อ') > -1);

// ── การต่อสายเข้า UI / workbook / ทะเบียน ────────────────────────────
assert(html.indexOf('id="xpDashboardBtn"') > -1, 'ต้องมีปุ่ม Dashboard Snapshot ในระดับ 1');
const initSrc = html.slice(html.indexOf('(function initExportCenter()'));
assert(/xpDashboardBtn'\);[\s\S]{0,120}xpRunDashboardSnapshot/.test(initSrc), 'ปุ่มต้องผูกกับ xpRunDashboardSnapshot');
const busySrc = grabFn('xpSetBusy');
assert(/xpDashboardBtn/.test(busySrc), 'ปุ่ม Dashboard ต้องถูกปิดระหว่างกำลังทำงาน กันกดซ้ำ');

const wbSrc = grabFn('xpRunWorkbook');
assert(/xpDashboardData\(\)/.test(wbSrc), 'Executive Workbook ต้องดึงข้อมูล Dashboard');
assert(/var wantDash = keys\.indexOf\('dashboardKpi'\) > -1;/.test(wbSrc),
  'ต้องใส่ชีต Dashboard เฉพาะเมื่อผู้ใช้ติ๊กชุด dashboardKpi ไว้ ไม่งั้น checkbox เหมือนไม่ทำงาน');
assert(/xpDashSheets\(dash\)/.test(wbSrc), 'Executive Workbook ต้องมีชีต Dashboard');
assert(/xpDashChartSeries\(dash\)/.test(wbSrc), 'ชีตกราฟต้องรวมกราฟของ Dashboard ด้วย');
assert(/return null;/.test(wbSrc), 'ถ้าดึง Dashboard ไม่ได้ ต้องข้ามไป ไม่ล้มทั้งไฟล์');
assert(/r\.ds\.key === 'dashboardKpi'\) return;/.test(wbSrc), 'ชุด dashboardKpi ต้องไม่ถูกใส่ซ้ำสองชีต');

const pdfSrc = grabFn('xpRunDashboardSnapshot');
assert(/window\.open/.test(pdfSrc), 'Dashboard Snapshot ต้องเปิดหน้าพรินต์');
assert(/xpSvgBarChart|xpSvgHBarChart/.test(pdfSrc), 'หน้าพรินต์ต้องมีกราฟ');
assert(!/canvas/.test(pdfSrc), 'หน้าพรินต์ต้องใช้ SVG ไม่ใช่ canvas (canvas พรินต์ไม่ติดในบางเบราว์เซอร์)');
assert(/stroke-dashoffset/.test(pdfSrc), 'ต้องมีเกจคะแนนสุขภาพ (วงกลมความคืบหน้า)');
assert(/Factory Command Center/.test(pdfSrc), 'หน้าพรินต์ต้องใช้ชื่อเดียวกับหน้า Dashboard');
assert(/xpLogExport\('dashboard-snapshot'/.test(pdfSrc), 'ต้องขึ้นทะเบียนการส่งออกด้วย');
assert(/ภาพรวมสถานะโรงงาน ณ เวลาที่สร้างไฟล์/.test(pdfSrc),
  'ต้องเขียนกำกับว่าเป็นสถานะปัจจุบัน ไม่ใช่ยอดของช่วงวันที่ที่เลือก — ไม่งั้นคนอ่านเข้าใจผิด');

// ชุดข้อมูลใน ระดับ 2
const dsSrc = grabVar('XP_DATASETS');
assert(/key: 'dashboardKpi'/.test(dsSrc), 'ต้องมีชุดข้อมูล dashboardKpi ให้เลือกในระดับ 2');
assert(/xpDashboardData\(\)\.then\(xpDashKpiTable\)/.test(dsSrc), 'ชุด dashboardKpi ต้องใช้ตัวคำนวณเดียวกัน');

console.log('export-dashboard-snapshot: OK (KPI ' + Object.keys(kpi).length + ' การ์ด, กราฟ ' + dashCharts.length + ' ตัว, ชีต ' + sheets.length + ' ใบ)');
