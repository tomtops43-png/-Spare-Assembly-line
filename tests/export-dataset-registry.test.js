// ทะเบียนชุดข้อมูลของ Export Center ต้องครบและต้องไม่ปล่อยของที่ห้ามหลุด
// เหตุผลที่ต้องมีเทสต์นี้: ไฟล์จากหน้านี้ออกไปอยู่นอกองค์กร (มือลูกค้า/ผู้จัดการ)
// ถ้าวันหลังมีคนเพิ่มชุดข้อมูลใหม่แล้วเผลอใส่คอลัมน์รหัสผ่าน/โทเคนเข้าไป จะไม่มีใครรู้เลย
const { html, backend, grabVar, buildModule, assert } = require('./_export-extract');

const XP_DATASETS = buildModule([grabVar('XP_DATASETS')], 'XP_DATASETS');

// ── ครบทุกชุดที่ประกาศไว้ในแผน ────────────────────────────────────────
assert(XP_DATASETS.length >= 20, 'ต้องมีชุดข้อมูลอย่างน้อย 20 ชุด แต่มี ' + XP_DATASETS.length);

const keys = XP_DATASETS.map(function(d) { return d.key; });
assert.strictEqual(new Set(keys).size, keys.length, 'key ของชุดข้อมูลต้องไม่ซ้ำกัน');

const required = [
  'stockMaster', 'logs', 'partTags', 'partTagGroups', 'machines', 'itemAudit',
  'partTagGroupItems', 'orderRequests', 'prHeaders', 'prLines', 'prAudit', 'purchaseHistory',
  'purchaseHistoryLog', 'purchaseImportLog', 'miscExpenses', 'miscExpenseLog',
  'productionVolume', 'productionCostConfig', 'stockCount', 'stockCountItems',
  'anomalies', 'userRoster', 'exportLog'
];
required.forEach(function(k) {
  assert(keys.indexOf(k) > -1, 'ขาดชุดข้อมูล: ' + k);
});

XP_DATASETS.forEach(function(d) {
  assert(d.name && typeof d.name === 'string', d.key + ' ต้องมีชื่อภาษาคน');
  assert(d.meta && typeof d.meta === 'string', d.key + ' ต้องมีคำอธิบายว่าได้อะไร');
  assert(d.group && typeof d.group === 'string', d.key + ' ต้องอยู่ในกลุ่ม');
  assert(typeof d.fetch === 'function', d.key + ' ต้องมีตัวดึงข้อมูล');
});

// กลุ่มต้องตรงกับหัวข้อบนหน้าจอ ไม่งั้นจะมีหัวข้อกลุ่มโผล่มาโดยไม่ได้ออกแบบ
const allowedGroups = ['คลังอะไหล่', 'การสั่งซื้อ', 'ต้นทุนและการผลิต', 'การควบคุมและตรวจสอบ'];
XP_DATASETS.forEach(function(d) {
  assert(allowedGroups.indexOf(d.group) > -1, d.key + ' อยู่กลุ่มที่ไม่รู้จัก: ' + d.group);
});

// ── ชุดที่ต้องจำกัดเฉพาะ Admin ────────────────────────────────────────
['userRoster', 'exportLog'].forEach(function(k) {
  const ds = XP_DATASETS.filter(function(d) { return d.key === k; })[0];
  assert(ds.adminOnly === true, k + ' ต้องเป็น adminOnly — ข้อมูลผู้ใช้/ทะเบียนการส่งออกไม่ใช่ของทุกคน');
});

// ── ชุดหลักต้องถูกติ๊กไว้ให้ตั้งแต่เปิดหน้า (กดปุ่มเดียวได้ของที่จำเป็น) ──
const core = XP_DATASETS.filter(function(d) { return d.core; }).map(function(d) { return d.key; });
['stockMaster', 'logs', 'orderRequests', 'purchaseHistory', 'miscExpenses', 'stockCount'].forEach(function(k) {
  assert(core.indexOf(k) > -1, k + ' ควรเป็นชุดหลัก (core) เพราะเป็นข้อมูลที่ผู้จัดการถามหาเสมอ');
});

// ── ห้ามมีทางหลุดของรหัสผ่าน / โทเคน ─────────────────────────────────
// 1) ฝั่งเว็บ: ไม่มีชุดข้อมูลไหนอ่านชีต Users ตรงๆ
const datasetSrc = grabVar('XP_DATASETS');
assert(datasetSrc.indexOf('listUsers') === -1, 'ห้ามใช้ listUsers ในทะเบียนชุดข้อมูล — มันคืนข้อมูลผู้ใช้ทั้งก้อน');
assert(/exportUserRoster/.test(datasetSrc), 'ชุดผู้ใช้ต้องผ่าน exportUserRoster ที่ตัดรหัสผ่านฝั่งเซิร์ฟเวอร์แล้ว');

// 2) ฝั่งเซิร์ฟเวอร์: whitelist ชีตดิบต้องไม่มีชีตที่เก็บรหัสผ่าน/โทเคน
const whitelist = buildModule([grabVar('EXPORT_RAW_SHEET_WHITELIST', backend)], 'EXPORT_RAW_SHEET_WHITELIST');
assert(whitelist.length > 0, 'whitelist ต้องไม่ว่าง');
['Users', 'users'].forEach(function(bad) {
  assert(whitelist.indexOf(bad) === -1, 'whitelist ห้ามมีชีต ' + bad + ' — มีคอลัมน์ password และ session_token');
});

// 3) exportUserRoster ต้องบังคับ Admin และต้องไม่ส่ง password / token ออกไป
const rosterSrc = backend.slice(backend.indexOf('function exportUserRoster('), backend.indexOf('function exportManifest('));
assert(/requireAdminUser/.test(rosterSrc), 'exportUserRoster ต้องบังคับสิทธิ์ Admin');
assert(!/u\.password/.test(rosterSrc), 'exportUserRoster ห้ามส่ง password ออกไป');
assert(!/u\.token/.test(rosterSrc), 'exportUserRoster ห้ามส่ง session token ออกไป');
assert(/allowed\.sort\(\)\.join/.test(rosterSrc), 'exportUserRoster ควรส่งเฉพาะรายชื่อสิทธิ์ที่เปิด');

// 4) dumpSheetForExport ต้องรองรับการตัดคอลัมน์ตั้งแต่ฝั่งเซิร์ฟเวอร์
const dumpSrc = backend.slice(backend.indexOf('function dumpSheetForExport('), backend.indexOf('function exportPrBundle('));
assert(/dropColumns/.test(dumpSrc), 'dumpSheetForExport ต้องมี dropColumns เพื่อกันคอลัมน์ที่ห้ามหลุด');
// Date object ที่ปล่อยผ่าน JSON.stringify จะกลายเป็น UTC แล้ววันเพี้ยนไป 1 วัน
assert(/Asia\/Bangkok/.test(dumpSrc), 'dumpSheetForExport ต้องแปลงวันที่ที่โซนเวลาไทย');

// ── ทุกชุดที่บอกว่านับแถวจากชีตไหน ต้องอ้างชื่อชีตที่มีอยู่จริงในระบบ ──
const knownSheets = [
  'Log', 'OrderRequests', 'PRHeaders', 'PRLines', 'PRAudit', 'PurchaseHistory',
  'PurchaseHistoryLog', 'PurchaseHistoryImportLog', 'MiscExpenses', 'MiscExpenseLog',
  'ProductionVolume', 'ProductionCostConfig', 'Machines', 'ItemAudit', 'PartTags',
  'PartTagGroups', 'PartTagGroupItems', 'StockCount', 'AnomalyAlerts', 'ExportLog',
  'Main List Stock', 'Stock for MC', 'Standard Spare part', 'Arc chut', 'Common Gv.2',
  'Gv.2 (6 plate)', 'Gv.2 (9 plate)', 'Coil Winding', 'Lug&Screw'
];
XP_DATASETS.forEach(function(d) {
  (d.countSheets || []).forEach(function(name) {
    assert(knownSheets.indexOf(name) > -1, d.key + ' อ้างชีตที่ไม่รู้จัก: ' + name);
  });
});

// ── ปุ่ม/ช่องบนหน้าจอที่เอนจินอ้างถึง ต้องมีอยู่ใน markup จริง ─────────
[
  'exportPage', 'xpDateFrom', 'xpDateTo', 'xpLineFilter', 'xpCategoryFilter',
  'xpTxnTypeFilter', 'xpGroupBy', 'xpCustomerMode', 'xpQuickRange', 'xpFilterSummary',
  'xpDatasetGrid', 'xpReportGrid', 'xpReportPreview', 'xpWorkbookBtn', 'xpPdfBtn',
  'xpJsonBundleBtn', 'xpExportXlsxBtn', 'xpExportCsvBtn', 'xpExportJsonBtn',
  'xpProgressWrap', 'xpProgressBar', 'xpProgressText', 'xpProgressLog',
  'xpSelectAllBtn', 'xpSelectNoneBtn', 'xpSelectCoreBtn', 'xpSelectedCount',
  'xpManifestBadge', 'xpRefreshManifestBtn', 'xpAuditCard', 'xpAuditRefreshBtn',
  'xpAuditTableWrap', 'tabExport'
].forEach(function(id) {
  assert(html.indexOf('id="' + id + '"') > -1, 'ไม่พบ element id="' + id + '" ใน index.html');
});

// ปุ่มช่วงเวลาแบบกดเร็วต้องมีครบทุกตัวเลือกที่ xpApplyQuickRange รองรับ
['today', '7d', 'month', 'prevmonth', 'quarter', 'year', 'all'].forEach(function(k) {
  assert(html.indexOf('data-xp-range="' + k + '"') > -1, 'ขาดปุ่มช่วงเวลา: ' + k);
});

console.log('export-dataset-registry: OK (' + XP_DATASETS.length + ' ชุดข้อมูล)');
