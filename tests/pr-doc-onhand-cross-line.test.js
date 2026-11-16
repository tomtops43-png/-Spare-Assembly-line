const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

// ปริ้น PR จาก Inbox: ไลน์ของทุกรายการ = ไลน์หัว PR (เช่น Arc Chute) แต่อะไหล่จริงอยู่ Lug&Screw
// เลข no เป็นเลขรันของแต่ละไลน์ → ห้ามดึง On hand ของ "อะไหล่คนละตัวที่ no ซ้ำกัน" ใน Arc Chute มาใส่
function grab(name) {
  const start = html.indexOf(name);
  assert(start > -1, 'cannot find ' + name);
  let i = html.indexOf('{', start), depth = 0;
  for (; i < html.length; i += 1) { if (html[i] === '{') depth += 1; else if (html[i] === '}' && --depth === 0) break; }
  return html.slice(start, i + 1);
}
const consts = html.match(/var PR_DOC_LOGO_PATH = [^\n]+\n/)[0] + html.match(/var PR_DOC_EN_MONTHS = [^\n]+\n/)[0];
const build = new Function('partsData', 'dashboardAllPartsData',
  'var window = { location: { href: "https://example.com/app/index.html" } };\n' +
  'var document = { getElementById: function() { return null; } };\n' +
  'function canonicalLineName(l) { return String(l || "-").trim() || "-"; }\n' +
  grab('function firstDefined(values, fallbackValue)') + '\n' + consts +
  grab('function prDocLogoUrl()') + '\n' + grab('function prDocStockStatus(item)') + '\n' +
  grab('function buildDocumentGradeReport(selected, options)') + '\n' +
  'return buildDocumentGradeReport;'
);
const master = [
  // อะไหล่ Arc Chute ที่บังเอิญ no ตรงกับของ Lug&Screw
  { no: '12', line: 'Arc Chute', name: 'Other part A', model: 'XYZ-1', stock: 4, min: 2, max: 4 },
  { no: '13', line: 'Arc Chute', name: 'Other part B', model: 'XYZ-2', stock: 2, min: 1, max: 3 },
  // อะไหล่จริง อยู่ไลน์ Lug&Screw หมดสต็อก
  { no: '12', line: 'Lug&Screw', name: 'Proximity sensor', model: 'SME-8M-DS-24V-K-2,5-OE', stock: 0, min: 1, max: 2 },
  { no: '13', line: 'Lug&Screw', name: 'cyl, compact slide table', model: 'MXH10-30Z', stock: 0, min: 1, max: 2 },
  { no: '20', line: 'Lug&Screw', name: 'Solenoid Valve', model: 'SY3120-5LZD-C4', stock: 0, min: 1, max: 2 },
  // อะไหล่ Arc Chute จริง (no+model ตรง) ต้องยังจับคู่ได้ตามเดิม
  { no: '5', line: 'Arc Chute', name: 'cyl, compact', model: 'CDQ2A32-10DZ', stock: 0, min: 1, max: 3 }
];
const items = [
  { no: '5', line: 'Arc Chute', name: 'cyl, compact', model: 'CDQ2A32-10DZ', suggestedQty: 1, unit_price: 1477.75 },
  { no: '12', line: 'Arc Chute', name: 'Proximity sensor', model: 'SME-8M-DS-24V-K-2,5-OE', suggestedQty: 2, unit_price: 1501.2 },
  { no: '13', line: 'Arc Chute', name: 'cyl, compact slide table', model: 'MXH10-30Z', suggestedQty: 2, unit_price: 4991.25 },
  { no: '20', line: 'Arc Chute', name: 'Solenoid Valve', model: 'SY3120-5LZD-C4', suggestedQty: 2, unit_price: 1563.75 }
];
const doc = build([], master)(items, { prNumber: 'PR-T' });
assert(doc.includes('Out of Stock: <b>4</b>'), 'ทั้ง 4 รายการหมดสต็อกจริง (ต้องไม่ดึง stock ของอะไหล่ Arc Chute ที่ no ซ้ำ)');
assert(doc.includes('In Stock: <b>0</b>'), 'ห้ามขึ้น In Stock');
assert(!doc.includes('pill-na'), 'Solenoid ต้องหาเจอข้ามไลน์ด้วยรุ่น ไม่ขึ้น N/A');
assert(!doc.includes('exceeds Max'), 'Max ต้องมาจากอะไหล่ตัวจริง (0 + 2 ไม่เกิน 2)');
console.log('pr-doc-onhand-cross-line: ok');
