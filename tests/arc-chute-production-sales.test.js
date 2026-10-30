// ยอดขาย (มูลค่าผลิต) ไลน์ Arc Chute: ดึงจาก ProductionLog ของไฟล์ Assembly-M-C-Management
// แล้วคิดราคา — ชีตราคากลางก่อน ถ้าไม่มีรุ่นให้ใช้ UnitPrice ในชีต Products ของไฟล์ต้นทางเอง
const { backend, grabFn, assert } = require('./_export-extract');

const SRC_ID = '1tz5Unlu2W5Zz5AUCfD1vBTdyj-FPLzOpEaUibOyhcfM';
const PRICE_ID = 'central-price';
const files = {};
files[SRC_ID] = {
  ProductionLog: [
    ['LogID', 'Timestamp', 'Date', 'MachineID', 'ProductCode', 'PlannedQty', 'ActualQty', 'DefectQty'],
    ['a', '', new Date('2026-10-05T00:00:00+07:00'), 'AS-MED', 'S1A23976-D', 750, 3500, 0],
    ['b', '', new Date('2026-10-08T00:00:00+07:00'), 'AC-07', '51207116JR', 0, -12, 12], // แถวงานเสีย = ข้าม
    ['c', '', new Date('2026-10-09T00:00:00+07:00'), 'AC-07', '51207116JR', 2000, 18000, 0],
    ['d', '', new Date('2026-11-01T00:00:00+07:00'), 'GV-2', 'W813890060190-JR', 4000, 1000, 0]
  ],
  Products: [
    ['ProductCode', 'ProductName', 'DefaultQty', 'Active', 'UnitPrice', 'Capacity'],
    ['S1A23976-D', 'Arc stack Medium', 0, true, 9.5, 750],
    ['51207116JR', 'Arc chute', 0, true, 2, 2000]
  ]
};
files[PRICE_ID] = {
  'Unit Price Latest': [
    ['Specification (รหัสรุ่น)', 'Unit Price (Latest)'],
    ['51207116JR', 1.5] // ชีตกลางมีราคา = ใช้ราคากลาง ไม่ใช้ราคาในไฟล์ต้นทาง
  ]
};

function sheetOf(rows) { return { getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }) }; }
const SpreadsheetApp = {
  openById: id => ({ getSheetByName: n => (files[id] && files[id][n] ? sheetOf(files[id][n]) : null) })
};
const pad = n => String(n).padStart(2, '0');
const Utilities = { formatDate: d => { const t = new Date(d.getTime() + 7 * 3600e3); return t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1); } };
const CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
const Logger = { log: () => {} };

// ตัด object literal ของ productionLogSources จริงจาก Backend.gs (บล็อกนี้ไม่มีปีกกาในสตริง/คอมเมนต์)
const srcDecl = 'SPARE_APP_CONFIG.productionLogSources = SPARE_APP_CONFIG.productionLogSources || ';
const srcStart = backend.indexOf(srcDecl) + srcDecl.length;
assert(srcStart > srcDecl.length - 1 && backend[srcStart] === '{', 'ต้องหา productionLogSources เจอ');
let depth = 0, srcEnd = srcStart;
for (; srcEnd < backend.length; srcEnd += 1) {
  if (backend[srcEnd] === '{') depth += 1;
  if (backend[srcEnd] === '}' && --depth === 0) break;
}
const sources = backend.slice(srcStart, srcEnd + 1);
const api = new Function('SpreadsheetApp', 'Utilities', 'CacheService', 'Logger',
  'var SPARE_APP_CONFIG = { productionLogSources: ' + sources + ', unitPriceSource: { id: "' + PRICE_ID + '", sheet: "Unit Price Latest", specCol: "Specification (รหัสรุ่น)", priceCol: "Unit Price (Latest)" } };\n' +
  'function requirePermission() {}\n' +
  'function getProductionVolumeSheet() { return { getDataRange: function() { return { getValues: function() { return [["Month"]]; } }; } }; }\n' +
  ['normalizeProductionSourceConfig', 'getExternalProductionVolumeForLine', 'getUnitPriceMap', 'getSourceUnitPriceMap',
    'readUnitPriceMap', 'matchModelPrice', 'getProductionVolume'].map(n => grabFn(n, backend, 0)).join('\n') +
  '\nreturn { SPARE_APP_CONFIG: SPARE_APP_CONFIG, getProductionVolume: getProductionVolume };'
)(SpreadsheetApp, Utilities, CacheService, Logger);

// เก็บเฉพาะไลน์ Arc Chute (ไลน์อื่นไม่มีไฟล์จำลอง = คืน null เงียบๆ)
const rows = api.getProductionVolume({}).filter(r => r.line === 'Arc Chute');
const oct = rows.find(r => r.month === '2026-10');
const nov = rows.find(r => r.month === '2026-11');

assert(oct && oct.source === 'auto', 'ต.ค. ต้องมาจากชีตผลิตจริง');
assert.strictEqual(oct.actual_qty, 3500 + 18000, 'แถว ActualQty ติดลบ (บันทึกงานเสีย) ต้องไม่ถูกนับ');
assert.strictEqual(oct.production_value, 3500 * 9.5 + 18000 * 1.5,
  'S1A23976-D ไม่มีในชีตกลาง → ใช้ราคา Products ของไฟล์ต้นทาง / 51207116JR ใช้ราคาชีตกลาง');
assert.strictEqual(oct.unpriced_qty, 0);

assert.strictEqual(nov.actual_qty, 1000);
assert.strictEqual(nov.production_value, 0, 'รุ่นที่ไม่มีราคาทั้งสองชีตต้องไม่ถูกเดาราคา');
assert.deepStrictEqual(nov.unpriced_models, { 'W813890060190-JR': 1000 });

console.log('Arc Chute production sales checks passed');
