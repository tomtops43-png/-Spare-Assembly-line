// ภาษาของไฟล์ที่ส่งออก (ไทย / อังกฤษ)
// ผู้บริหารต้องเอาไฟล์ไปเสนอลูกค้าเป็นภาษาอังกฤษ — ไฟล์ที่แปลครึ่งเดียวแย่กว่าไฟล์ไทยล้วน
// เพราะเปิดในที่ประชุมแล้วดูไม่จบ เทสต์นี้จึงบังคับสองชั้น:
//  1) สแกนแบบคงที่: ทุกข้อความไทยที่ไหลลงไฟล์ต้องมีคำแปลใน XP_LANG_EN
//  2) รันจริงด้วย lang='en' + ข้อมูลตัวอย่างที่เป็นอังกฤษล้วน แล้วเช็คว่า
//     ไม่มีอักษรไทยเหลืออยู่ในหัวคอลัมน์/ค่าที่ระบบสร้าง/หมายเหตุเลยแม้แต่ตัวเดียว
// ชั้นที่ 2 คือชั้นที่จับของจริง — คนเพิ่มคอลัมน์ใหม่แล้วลืมใส่คำแปลจะติดที่นี่ทันที
const { html, grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

const THAI = /[฀-๿]/;

const XP_LANG_EN = buildModule([grabVar('XP_LANG_EN')], 'XP_LANG_EN');
const XP_DICTIONARY = buildModule([grabVar('XP_DICTIONARY')], 'XP_DICTIONARY');

// ── โครงสร้างตารางแปล ───────────────────────────────────────────────
const langKeys = Object.keys(XP_LANG_EN);
assert(langKeys.length > 250, 'ตารางแปลควรมีอย่างน้อย 250 รายการ แต่มี ' + langKeys.length);
langKeys.forEach(function(k) {
  const v = XP_LANG_EN[k];
  assert(typeof v === 'string' && v.length, 'คำแปลของ "' + k + '" ต้องไม่ว่าง');
  assert(!THAI.test(v), 'คำแปลของ "' + k + '" ยังมีอักษรไทยอยู่: ' + v);
  // เทมเพลตต้องมี placeholder ครบทั้งสองภาษา ไม่งั้นตัวเลขหายไปจากประโยคอังกฤษ
  (k.match(/\{(\w+)\}/g) || []).forEach(function(ph) {
    assert(v.indexOf(ph) > -1, 'คำแปลของ "' + k + '" ขาด placeholder ' + ph);
  });
});

// ── ชั้นที่ 1: สแกนข้อความไทยทุกตัวที่ไหลลงไฟล์ ─────────────────────
const region = (function() {
  const start = html.indexOf('// EXPORT CENTER — เครื่องยนต์');
  const end = html.indexOf('checkSession().then(function() {');
  assert(start > -1 && end > start, 'ต้องหาโค้ด Export Center เจอ');
  return html.slice(start, end);
})();

function scan(re) {
  const out = [];
  let m;
  while ((m = re.exec(region))) out.push(m[1]);
  return out;
}
const missing = [];
function requireTranslation(list, where) {
  list.filter(function(s) { return THAI.test(s); }).forEach(function(s) {
    if (XP_LANG_EN[s] === undefined) missing.push(where + ': ' + s);
  });
}
// หัวคอลัมน์ทุกชุด (ผ่าน xpTable แล้วลงเป็นหัวตารางในไฟล์)
requireTranslation(scan(/label: '([^']*)'/g), 'label');
// ชื่อชุดข้อมูล / ชื่อชีต / คำอธิบาย / ชื่อรายงาน
requireTranslation(scan(/\bname: '([^']*)'/g), 'name');
requireTranslation(scan(/sheetName: '([^']*)'/g), 'sheetName');
requireTranslation(scan(/meta: '([^']*)'/g), 'meta');
requireTranslation(scan(/desc: '([^']*)'/g), 'desc');
// ทุกข้อความที่ส่งเข้า xpT() ต้องมีคำแปล ไม่งั้นเรียกไปก็ได้ไทยกลับมา
requireTranslation(scan(/xpT\('([^']*)'/g), 'xpT()');
// ป้ายและหน่วยของ KPI (ส่งผ่าน add(group, label, value, unit, note))
requireTranslation(scan(/\badd\('([^']*)'/g), 'KPI group');
requireTranslation(scan(/\badd\('[^']*', '([^']*)'/g), 'KPI label');
// คำอธิบายในพจนานุกรมข้อมูล
requireTranslation(Object.keys(XP_DICTIONARY), 'dictionary key');
requireTranslation(Object.keys(XP_DICTIONARY).map(function(k) { return XP_DICTIONARY[k]; }), 'dictionary description');

assert.strictEqual(missing.length, 0,
  'ยังมีข้อความไทยที่ไม่มีคำแปล ' + missing.length + ' รายการ:\n  ' + missing.join('\n  '));

// ── ชั้นที่ 2: รันจริงเป็นภาษาอังกฤษแล้วต้องไม่มีอักษรไทยเหลือ ────────
const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "2026-09-01", to: "2026-09-10", line: "", category: "", txnType: "", groupBy: "month", lang: "th" };',
    'var currentUser = { username: "manager01", role: "leader" };',
    grabVar('XP_DICTIONARY'),
    grabVar('XP_REPORTS'),
    grabFn('xpBuildKpis'),
    grabFn('xpKpiTable'),
    grabFn('xpBuildTrend'),
    grabFn('xpBuildLineSummary'),
    grabFn('xpBuildDictionary'),
    grabFn('xpBuildCover'),
    grabFn('xpBuildAttachmentSheet'),
    grabFn('xpSafeSheetName'),
    grabFn('xpFileSuffix')
  ]),
  '{ XP_REPORTS: XP_REPORTS, xpBuildKpis: xpBuildKpis, xpKpiTable: xpKpiTable, xpBuildTrend: xpBuildTrend, xpBuildLineSummary: xpBuildLineSummary, xpBuildDictionary: xpBuildDictionary, xpBuildCover: xpBuildCover, xpBuildAttachmentSheet: xpBuildAttachmentSheet, xpFileSuffix: xpFileSuffix, xpT: xpT, setLang: function(l) { xpFilters.lang = l; } }'
);

// ข้อมูลตัวอย่างเป็นอังกฤษล้วน — อักษรไทยที่หลุดออกมาจึงมาจากโค้ดแน่นอน ไม่ใช่จากข้อมูล
const master = [
  { __sheet: 'LugScrew', no: '1', name: 'Breaker', model: 'BK-10', line: 'LugScrew', category: 'Electrical', location: 'A-01', stock: 4, min: 10, max: 40, unit: 'PCS', unit_price: 100, supplier: 'ABB', drawing_url: 'https://d/1' },
  { __sheet: 'LugScrew', no: '2', name: 'Bolt M8', model: 'M8', line: 'LugScrew', category: 'Hardware', location: 'A-02', stock: 200, min: 50, max: 300, unit: 'PCS', unit_price: 2 },
  { __sheet: 'H9', no: '3', name: 'Belt', model: 'B-77', line: 'H9', category: 'Mechanical', location: 'B-01', stock: 0, min: 2, max: 10, unit: 'PCS', unit_price: 500 },
  { __sheet: 'H9', no: '4', name: 'Old Fuse', model: 'FZ-1', line: 'H9', category: 'Electrical', location: 'B-02', stock: 20, min: 0, max: 5, unit: 'PCS', unit_price: 30 }
];
const logs = [
  { timestamp: '2026-09-02 08:00:00', type: 'Output', process: 'LugScrew', category: 'Electrical', partName: 'Breaker', model: 'BK-10', qty: 6, by: 'somchai', machine: 'LS-10', stockBefore: 10, stockAfter: 4 },
  { timestamp: '2026-09-03 08:00:00', type: 'Output', process: 'LugScrew', category: 'Hardware', partName: 'Bolt M8', model: 'M8', qty: 50, by: 'somchai', machine: 'LS-10', stockBefore: 250, stockAfter: 200 },
  { timestamp: '2026-09-04 08:00:00', type: 'Output', process: 'H9', category: 'Mechanical', partName: 'Belt', model: 'B-77', qty: 8, by: 'somsak', machine: 'H9-1', stockBefore: 8, stockAfter: 0 },
  { timestamp: '2026-09-05 08:00:00', type: 'Input', process: 'H9', category: 'Mechanical', partName: 'Belt', model: 'B-77', qty: 4, by: 'admin', machine: '', stockBefore: 0, stockAfter: 4 }
];
const ctx = {
  master: master,
  logs: logs,
  orderRequests: [{ request_id: 'RQ1', line: 'H9', item_name: 'Belt', model: 'B-77', status: 'Approved', requested_date: '2026-09-01', approved_date: '2026-09-04', request_qty: 2, unit_price: 500, unit: 'PCS' }],
  purchaseHistory: [
    { 'History ID': 'PH1', Line: 'H9', 'Part Name': 'Belt', 'Model / Part No.': 'B-77', Status: 'Received', 'Requested Date': '2026-09-01', 'Ordered Date': '2026-09-05', 'Received Date': '2026-09-09', 'Unit Price': 500, 'Qty Ordered': 2, 'Total Amount': 1000, Brand: 'Gates' },
    { 'History ID': 'PH2', Line: 'H9', 'Part Name': 'Belt', 'Model / Part No.': 'B-77', Status: 'Received', 'Requested Date': '2026-09-02', 'Ordered Date': '2026-09-06', 'Received Date': '2026-09-10', 'Unit Price': 650, 'Qty Ordered': 1, 'Total Amount': 650, Brand: 'Optibelt' }
  ],
  miscExpenses: [{ expense_id: 'EX1', month: '2026-09', date: '2026-09-03', line: 'LugScrew', category: 'Hardware', item_name: 'Wrench', total_amount: 300, receipt_url: 'https://d/r1' }],
  productionVolume: [{ month: '2026-09', line: 'LugScrew', actual_qty: 1000, production_value: 10000, source: 'auto' }],
  costConfig: [{ line: 'LugScrew', unit_price: 10, target_pct: 12 }],
  stockCount: [{ session_id: 'S1', month: '2026-09', line: 'H9', round_no: 1, total_items: 100, matched: 95, diff_count: 5, adjusted_count: 5, status: 'approved', created_by: 'somchai' }],
  partTags: [
    { tag_no: 'T1', part_name: 'Blade', model: 'BL-1', line: 'LugScrew', machine: 'LS-10', installed_at: '2026-01-01', removed_at: '2026-03-02', status: 'removed' },
    { tag_no: 'T2', part_name: 'Blade', model: 'BL-1', line: 'LugScrew', machine: 'LS-10', installed_at: '2026-03-02', removed_at: '2026-05-01', status: 'removed' }
  ],
  machines: [{ machine_id: 'MC-1', line: 'H9', machine_name: 'H9-1', active: true, created_by: 'admin' }],
  prBundle: { prHeaders: { headers: ['pr_id', 'status', 'total_amount'], rows: [['PR-1', 'PENDING', 5000]] }, prLines: { headers: ['pr_id'], rows: [['PR-1']] }, prAudit: { headers: ['pr_id'], rows: [] } },
  failed: {}
};

function findThai(table, where) {
  const hits = [];
  (table.headers || []).forEach(function(h, i) {
    if (THAI.test(String(h))) hits.push(where + ' header[' + i + ']: ' + h);
  });
  (table.rows || []).forEach(function(r, ri) {
    (r || []).forEach(function(c, ci) {
      if (typeof c === 'string' && THAI.test(c)) hits.push(where + ' row[' + ri + '][' + ci + ']: ' + c);
    });
  });
  return hits;
}

mod.setLang('en');
let leaks = [];

// KPI / เทรนด์ / สรุปตามไลน์
const kpis = mod.xpBuildKpis(ctx);
leaks = leaks.concat(findThai(mod.xpKpiTable(kpis), 'KPI'));
kpis.forEach(function(k) {
  if (THAI.test(String(k.group))) leaks.push('KPI group: ' + k.group);
  if (THAI.test(String(k.label))) leaks.push('KPI label: ' + k.label);
  if (THAI.test(String(k.unit))) leaks.push('KPI unit: ' + k.unit);
  if (THAI.test(String(k.note))) leaks.push('KPI note: ' + k.note);
  if (typeof k.value === 'string' && THAI.test(k.value)) leaks.push('KPI value: ' + k.value);
});
const trend = mod.xpBuildTrend(ctx);
leaks = leaks.concat(findThai(trend.table, 'Trend'));
leaks = leaks.concat(findThai(mod.xpBuildLineSummary(ctx), 'LineSummary'));

// รายงานวิเคราะห์ทั้ง 10 ตัว + ชีตแนบ + หมายเหตุ
mod.XP_REPORTS.forEach(function(rp) {
  const built = rp.build(ctx);
  leaks = leaks.concat(findThai(built.table, 'Report ' + rp.key));
  if (built.note && THAI.test(built.note)) leaks.push('Report ' + rp.key + ' note: ' + built.note);
  (built.extraSheets || []).forEach(function(sheet) {
    if (THAI.test(String(sheet.name))) leaks.push('Report ' + rp.key + ' extra sheet name: ' + sheet.name);
    leaks = leaks.concat(findThai(sheet.table, 'Report ' + rp.key + ' extra'));
  });
  if (THAI.test(mod.xpT(rp.name))) leaks.push('Report name: ' + rp.name);
  if (THAI.test(mod.xpT(rp.desc))) leaks.push('Report desc: ' + rp.desc);
});

// หน้าปก (ทั้งกรณีสำเร็จและกรณีมีชุดที่ดึงไม่ได้)
leaks = leaks.concat(findThai(mod.xpBuildCover([
  { ds: { name: 'Master อะไหล่', sheetName: 'Master อะไหล่' }, table: { headers: ['a'], rows: [[1]] } },
  { ds: { name: 'PR audit', sheetName: 'PR audit' }, error: 'HTTP 500' }
]), 'Cover'));

// ชีตลิงก์ไฟล์แนบ
leaks = leaks.concat(findThai(mod.xpBuildAttachmentSheet(ctx, []), 'Attachments'));

// พจนานุกรมข้อมูล — คำอธิบายต้องเป็นอังกฤษด้วย ไม่ใช่แปลแค่หัวคอลัมน์
// ใช้ตารางที่มีคอลัมน์ซึ่ง XP_DICTIONARY มีคำอธิบายไว้จริง (คงเหลือ / มูลค่าคงคลัง)
// เพื่อพิสูจน์ว่าการแปลกลับ (อังกฤษ → ไทย) หาคำอธิบายเจอ
const dictSheet = mod.xpBuildDictionary([
  { name: 'Master อะไหล่', table: mod.XP_REPORTS[1].build(ctx).table },
  { name: 'Log รับเข้า-เบิกออก', table: mod.xpBuildLineSummary(ctx) }
]);
leaks = leaks.concat(findThai(dictSheet, 'Dictionary'));

assert.strictEqual(leaks.length, 0,
  'มีภาษาไทยหลุดเข้าไฟล์ภาษาอังกฤษ ' + leaks.length + ' จุด:\n  ' + leaks.slice(0, 40).join('\n  '));

// ── พจนานุกรมยังต้องหาคำอธิบายเจอ ทั้งที่หัวคอลัมน์ถูกแปลไปแล้ว ──────
const dictMap = {};
dictSheet.rows.forEach(function(r) { dictMap[r[1]] = r[2]; });
assert(dictMap['On Hand'], 'ต้องมีบรรทัดของคอลัมน์ On Hand');
assert(dictMap['On Hand'].indexOf('Column taken directly') === -1,
  'คอลัมน์ On Hand มีคำอธิบายจริงอยู่แล้ว ต้องไม่ตกไปใช้ข้อความ fallback (แปลกลับหาคีย์ไม่เจอ)');
assert(/latest unit price/.test(dictMap['Inventory Value'] || ''), 'คำอธิบายมูลค่าคงคลังต้องถูกแปลและหาเจอ');

// ── โหมดไทยต้องไม่เปลี่ยนอะไรเลย ─────────────────────────────────────
mod.setLang('th');
const thTable = mod.xpBuildLineSummary(ctx);
assert(thTable.headers.indexOf('ไลน์') > -1, 'โหมดไทยต้องได้หัวคอลัมน์ภาษาไทยเหมือนเดิม');
assert(thTable.headers.indexOf('Line') === -1, 'โหมดไทยต้องไม่มีหัวคอลัมน์อังกฤษปนมา');
mod.setLang('en');
const enTable = mod.xpBuildLineSummary(ctx);
assert(enTable.headers.indexOf('Line') > -1, 'โหมดอังกฤษต้องได้ Line');
assert.strictEqual(enTable.headers.length, thTable.headers.length, 'จำนวนคอลัมน์ต้องเท่ากันทั้งสองภาษา');
assert.strictEqual(enTable.rows.length, thTable.rows.length, 'จำนวนแถวต้องเท่ากันทั้งสองภาษา');

// ── ชื่อไฟล์ต้องบอกได้ว่าเป็นไฟล์ภาษาอังกฤษ ───────────────────────────
assert(/EN/.test(mod.xpFileSuffix()), 'ไฟล์ภาษาอังกฤษต้องมี EN ในชื่อไฟล์ กันส่งไฟล์ผิดภาษาให้ลูกค้า');
mod.setLang('th');
assert(!/EN/.test(mod.xpFileSuffix()), 'ไฟล์ภาษาไทยต้องไม่มี EN ในชื่อไฟล์');

// ── UI: ช่องเลือกภาษาและการเดินสายค่า ─────────────────────────────────
assert(html.indexOf('id="xpLangFilter"') > -1, 'ต้องมีช่องเลือกภาษาในแถบตัวกรอง');
assert(/<option value="en">/.test(html), 'ต้องมีตัวเลือกภาษาอังกฤษ');
const readSrc = grabFn('xpReadFilterInputs');
assert(/xpFilters\.lang = \(lg && lg\.value === 'en'\) \? 'en' : 'th'/.test(readSrc), 'ต้องอ่านค่าภาษาจากช่องเลือก');
const initSrc = html.slice(html.indexOf('(function initExportCenter()'));
assert(/'xpLangFilter'\]\.forEach/.test(initSrc) || /'xpCustomerMode', 'xpLangFilter'/.test(initSrc),
  'เปลี่ยนภาษาแล้วต้อง trigger xpReadFilterInputs (ผูก event ไว้)');
// เปลี่ยนภาษาต้องล้าง cache เพราะตารางที่สร้างไว้เป็นภาษาเดิม
const afterSrc = grabFn('xpAfterFilterChange');
assert(/xpClearCache\(\)/.test(afterSrc), 'เปลี่ยนภาษาต้องล้าง cache ตารางที่สร้างไว้ภาษาเดิม');
// บรรทัดสรุปต้องบอกภาษาของไฟล์ให้ผู้ใช้เห็นก่อนกดดาวน์โหลด
const summarySrc = grabFn('xpRenderFilterSummary');
assert(/xpIsEn\(\)/.test(summarySrc), 'บรรทัดสรุปต้องบอกว่าไฟล์ที่จะได้เป็นภาษาอะไร');
// ทะเบียนการส่งออกต้องบันทึกภาษาไว้ด้วย (ไล่ได้ว่าไฟล์ที่ส่งลูกค้าเป็นภาษาไหน)
const logSrc = grabFn('xpLogExport');
assert(/lang=/.test(logSrc), 'ทะเบียนการส่งออกต้องบันทึกภาษาของไฟล์');

// ── ค่าที่ยังไม่มีคำแปลต้องคืนข้อความไทยเดิม ไม่ใช่ค่าว่าง ─────────────
mod.setLang('en');
assert.strictEqual(mod.xpT('ข้อความที่ยังไม่มีคำแปลแน่ๆ'), 'ข้อความที่ยังไม่มีคำแปลแน่ๆ',
  'ไม่มีคำแปลต้องคืนข้อความเดิม (ผิดแบบเห็นได้ ดีกว่าช่องหาย)');
assert.strictEqual(mod.xpT('คงเหลือ'), 'On Hand');
assert.strictEqual(mod.xpT('เกินเป้า {n} จุด', { n: 3.5 }), 'Over target by 3.5 points', 'เทมเพลตต้องแทนค่าได้');
mod.setLang('th');
assert.strictEqual(mod.xpT('คงเหลือ'), 'คงเหลือ', 'โหมดไทยต้องคืนข้อความไทย');
assert.strictEqual(mod.xpT('เกินเป้า {n} จุด', { n: 3.5 }), 'เกินเป้า 3.5 จุด');

console.log('export-language: OK (' + langKeys.length + ' คำแปล, ไม่มีภาษาไทยหลุดเข้าไฟล์อังกฤษ)');
