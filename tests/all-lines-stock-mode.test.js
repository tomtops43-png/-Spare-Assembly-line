const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── "ทุกไลน์" ต้องเป็นตัวเลือกแรกในเมนูเลือกไลน์ ────────────────────────────
assert(/var lineConfig = \[\n\s*\{ key: ALL_LINES_KEY,/.test(htmlLf),
  '"ทุกไลน์" ต้องอยู่บนสุดของเมนูไลน์');
assert(htmlLf.includes("var ALL_LINES_LABEL = 'ทุกไลน์';"), 'ต้องมีป้ายชื่อภาษาไทยของไลน์เสมือน');
assert(/iconMap\[ALL_LINES_KEY\]/.test(htmlLf) && /subtitleMap\[ALL_LINES_KEY\]/.test(htmlLf),
  'การ์ดเมนูของ "ทุกไลน์" ต้องมีไอคอน/คำบรรยายของตัวเอง');
// กดแล้ววิ่งเข้าเส้นทางเดิม (โหลดข้อมูล + ทาสีธีม) เหมือนไลน์อื่น ไม่ต้องมีปุ่มพิเศษ
assert(/setCurrentLine\(btn\.getAttribute\('data-line'\)\);/.test(htmlLf),
  'ปุ่มเลือกไลน์ต้องผ่าน setCurrentLine เพื่อจำไลน์จริงล่าสุดไว้');

// ── หน้า Stock: โหมดทุกไลน์ต้องโหลดของทุกไลน์มารวมกัน ────────────────────────
assert(/if \(isAllLinesMode\(\)\) return loadAllLinesPartsData\(opts\);/.test(htmlLf),
  'loadPartsData ต้องแยกไปเส้นทางรวมทุกไลน์');
const loader = grab(/^ {4}function loadAllLinesPartsData\(opts\) \{[\s\S]*?\n {4}\}/m, 'loadAllLinesPartsData');
assert(loader.includes("loadLineCache(lineKey, lineKey + '::ALL')"), 'ต้องใช้แคชรายไลน์ก้อนเดิมของหน้า Stock');
assert(loader.includes("saveLineCache(lineKey, rows, lineKey + '::ALL')"), 'โหลดเสร็จต้องเขียนกลับแคชรายไลน์');
// ห้ามสร้างแคชรวมก้อนใหม่ — ซ้ำซ้อนกับแคชรายไลน์ และเสี่ยงชน quota ของ localStorage
assert(!/saveLineCache\(ALL_LINES_KEY/.test(htmlLf), 'ห้ามเก็บแคชรวมของไลน์เสมือน');
assert(loader.includes('applyLoadedPartsData(rows)'), 'ต้องส่งของเข้าหน้า Stock ผ่านทางเดิม');
assert(/chain = chain\.then/.test(loader), 'ต้องยิงทีละไลน์ ไม่ยิงพร้อมกันทุกไลน์');
assert(loader.includes('{ lite: true'), 'ต้องโหลดแบบ lite เหมือนโหมดไลน์เดียว');
assert(loader.indexOf('if (mergedRows().length) paint();') > -1, 'ของที่มีในแคชต้องขึ้นก่อน ไม่ปล่อยจอว่าง');

// สลับออกจากโหมดทุกไลน์ ห้ามเอาก้อนรวมไปทับแคชของไลน์ที่เพิ่งเลือก
assert(/loadedStockSheetKey\.indexOf\(ALL_LINES_KEY\) !== 0/.test(htmlLf),
  'ก้อนข้อมูลรวมทุกไลน์ต้องไม่ถูกเขียนกลับเป็นแคชของไลน์เดียว');
assert(/if \(!isAllLinesMode\(\)\) saveLineCache\(currentLine, partsData, getCurrentSheetName\(\)\);/.test(htmlLf),
  'รีเฟรชรายตัวหลังเบิก/รับ ก็ห้ามเขียนก้อนรวมทับแคชรายไลน์');

// ── การ์ดต้องบอกว่าเป็นของไลน์ไหน (เฉพาะโหมดทุกไลน์) ────────────────────────
const badge = grab(/^ {4}function renderCardLineBadge\(item, compact\) \{[\s\S]*?\n {4}\}/m, 'renderCardLineBadge');
assert(badge.indexOf("if (!isAllLinesMode()) return ''") > -1, 'โหมดไลน์เดียวไม่ต้องขึ้นป้ายไลน์ (รู้อยู่แล้ว)');
assert((htmlLf.match(/renderCardLineBadge\(item, (false|true)\) \+/g) || []).length === 2,
  'ต้องติดป้ายไลน์ทั้งการ์ดจอใหญ่และการ์ดมือถือ');

// ── ตัวกรองไลน์บนหน้า Stock ────────────────────────────────────────────────
assert(htmlLf.includes('<select id="lineFilter"') && htmlLf.includes('<select id="lineFilterMobile"'),
  'ต้องมีตัวกรองไลน์ทั้งจอใหญ่และมือถือ');
assert(/function populateLineFilterOptions\(data\)/.test(htmlLf), 'ต้องมีตัวเติมตัวเลือกไลน์');
assert(/ {6}populateLineFilterOptions\(partsData\);\n {6}populateCategoryOptions\(partsData\);/.test(htmlLf),
  'ต้องซิงก์ทุกครั้งที่ข้อมูลเปลี่ยน (สลับกลับไลน์เดียวแล้วต้องซ่อนเอง)');
const filters = grab(/^ {4}function applyFilters\(resetPage\) \{[\s\S]*?var filtered = partsData\.filter/m, 'applyFilters head');
assert(filters.indexOf("(lineFilterEl && isAllLinesMode()) ? lineFilterEl.value : 'all'") > -1,
  'ตัวกรองไลน์ต้องมีผลเฉพาะโหมดทุกไลน์');
assert(/if \(needLine && lineOfItem\(item\) !== selectedLineFilter\) return false;/.test(htmlLf),
  'ต้องกรองตามไลน์จริงของอะไหล่');

// ── รันของจริง: คีย์ไลน์เสมือนห้ามหลุดออกไปกับข้อมูลที่เขียนลงชีต ────────────
const src = [
  grab(/^ {4}var LINE_SHEET_OPTIONS = \{[\s\S]*?\n {4}\};/m, 'LINE_SHEET_OPTIONS'),
  grab(/^ {4}var ALL_LINES_KEY = '__ALL_LINES__';[\s\S]*?\n(?= {4}var LINE_CACHE_VERSION)/m, 'all-lines helpers'),
  grab(/^ {4}var currentLine = 'H9';[\s\S]*?\n {4}function formLineDefault\(item\) \{[\s\S]*?\n {4}\}/m, 'currentLine helpers'),
  grab(/^ {4}function normalizeLineText\(value\) \{[\s\S]*?\n {4}\}/m, 'normalizeLineText'),
  grab(/^ {4}function mapLine\(lineName\) \{[\s\S]*?\n {4}\}/m, 'mapLine'),
  grab(/^ {4}function canonicalLineName\(value\) \{[\s\S]*?\n {4}\}/m, 'canonicalLineName')
].join('\n');
const api = new Function(src + '\nreturn { ALL_LINES_KEY: ALL_LINES_KEY, LINE_SHEET_OPTIONS: LINE_SHEET_OPTIONS,' +
  ' realLineKeys: realLineKeys, lineOfItem: lineOfItem, lineDisplayName: lineDisplayName,' +
  ' canonicalLineName: canonicalLineName, isAllLinesMode: isAllLinesMode, setCurrentLine: setCurrentLine,' +
  ' activeRealLine: activeRealLine, formLineDefault: formLineDefault };')();

// ชีตของไลน์เสมือน = ชีตของทุกไลน์รวมกัน ไม่ซ้ำ
const union = api.LINE_SHEET_OPTIONS[api.ALL_LINES_KEY];
const realSheets = [];
api.realLineKeys().forEach(function(k) {
  api.LINE_SHEET_OPTIONS[k].forEach(function(s) { if (realSheets.indexOf(s) === -1) realSheets.push(s); });
});
assert.deepStrictEqual(union.slice().sort(), realSheets.slice().sort(), 'ไลน์เสมือนต้องครอบคลุมทุกชีตของทุกไลน์');
assert.strictEqual(new Set(union).size, union.length, 'ชีตต้องไม่ซ้ำ');
assert(api.realLineKeys().indexOf(api.ALL_LINES_KEY) === -1, 'รายชื่อไลน์จริงต้องไม่มีคีย์เสมือน');
assert.strictEqual(api.realLineKeys().length, 4, 'ต้องมีไลน์จริงครบ 4 ไลน์');

// ระบุไลน์ของอะไหล่จากชีตต้นทาง — เชื่อถือได้กว่าคอลัมน์ line ที่บางแถวว่าง
assert.strictEqual(api.lineOfItem({ __sourceSheet: 'Arc chut' }), 'Arc Chute');
assert.strictEqual(api.lineOfItem({ __sourceSheet: 'Gv.2 (9 plate)' }), 'Arc Chute');
assert.strictEqual(api.lineOfItem({ __sourceSheet: 'Stock for MC' }), 'H9');
assert.strictEqual(api.lineOfItem({ subLine: 'Coil Winding' }), 'Coil Winding');
assert.strictEqual(api.lineOfItem({ __sourceSheet: 'Lug&Screw', line: 'H9' }), 'Lug&Screw');

api.setCurrentLine(api.ALL_LINES_KEY);
assert(api.isAllLinesMode(), 'เลือกไลน์เสมือนแล้วต้องเข้าโหมดทุกไลน์');
assert.strictEqual(api.lineDisplayName(api.ALL_LINES_KEY), 'ทุกไลน์', 'ต้องโชว์ชื่อไทย ไม่ใช่คีย์ดิบ');
// คีย์เสมือนห้ามหลุดไปติดกับข้อมูลที่เขียนลงชีต (ใบ PR / log เบิก / ของติดดาว)
assert.notStrictEqual(api.canonicalLineName(''), api.ALL_LINES_KEY, 'canonicalLineName ห้ามคืนคีย์เสมือน');
assert.notStrictEqual(api.activeRealLine(), api.ALL_LINES_KEY, 'ไลน์สำหรับฟอร์มต้องเป็นไลน์จริงเสมอ');
assert.strictEqual(api.activeRealLine(), 'H9', 'ยังไม่เคยเลือกไลน์อื่น ต้องถอยไปไลน์เริ่มต้น');
assert.strictEqual(api.formLineDefault({ __sourceSheet: 'Arc chut' }), 'Arc Chute',
  'ฟอร์มต้องยึดไลน์ของอะไหล่ชิ้นนั้นก่อน');
assert.strictEqual(api.formLineDefault(null), 'H9', 'ไม่มีอะไหล่อ้างอิง ค่อยถอยไปไลน์จริงล่าสุด');

// เลือกไลน์จริงไปก่อนแล้วค่อยกลับมาโหมดทุกไลน์ — ต้องจำไลน์จริงล่าสุดไว้
api.setCurrentLine('Coil Winding');
assert(!api.isAllLinesMode());
api.setCurrentLine(api.ALL_LINES_KEY);
assert.strictEqual(api.activeRealLine(), 'Coil Winding', 'ต้องจำไลน์จริงล่าสุดที่ผู้ใช้ทำงานอยู่');
assert.strictEqual(api.formLineDefault(null), 'Coil Winding');

console.log('All-lines stock mode checks passed');
