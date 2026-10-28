const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(src, name) {
  const m = src.match(new RegExp('^ {0,4}function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n {0,4}\\}', 'm'));
  assert(m, 'ต้องดึง ' + name + ' ออกมาได้');
  return m[0];
}
// ฟังก์ชันระดับบนสุดของ Backend.gs (ปิดด้วย } ที่คอลัมน์แรก — ข้างในมี "  });" ซ้อนอยู่)
function grabTop(src, name) {
  const m = src.match(new RegExp('^function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n\\}', 'm'));
  assert(m, 'ต้องดึง ' + name + ' ออกมาได้');
  return m[0];
}
const planBlock = backend.match(/^var LOCATION_PLAN_BY_LINE = \{[\s\S]*?\n\};/m);
assert(planBlock, 'ต้องมีตาราง LOCATION_PLAN_BY_LINE');

const be = new Function([
  planBlock[0],
  grab(backend, 'normalizeLocationText'),
  grabTop(backend, 'locationPlanCategoryKey'),
  grabTop(backend, 'planLocationsForRows'),
  'return { LOCATION_PLAN_BY_LINE, locationPlanCategoryKey, planLocationsForRows };'
].join('\n'))();

// ── ตาราง ไลน์ × หมวด → Location ตามที่หน้างานจัดชั้นไว้ ──────────────────────────
const P = be.LOCATION_PLAN_BY_LINE;
assert.deepStrictEqual(P['H9'].locations, { pneumatic: 'SP-H9-A4-01', electrical: 'SP-H9-A4-02', mechanical: 'SP-H9-A4-03' });
assert.deepStrictEqual(P['Arc Chute'].locations, { electrical: 'SP-AC-A3-01', pneumatic: 'SP-AC-A3-02', mechanical: 'SP-AC-A3-03' });
assert.deepStrictEqual(P['Lug&Screw'].locations, { electrical: 'SP-LS-A2-01', pneumatic: 'SP-LS-A2-02', mechanical: 'SP-LS-A2-03' });
assert.deepStrictEqual(P['Coil Winding'].locations, { mechanical: 'SP-CL-A1-01', electrical: 'SP-CL-A1-02', pneumatic: 'SP-CL-A1-03' });

// ชีตของแต่ละไลน์ต้องตรงกับหน้าเว็บ ไม่งั้นอะไหล่บางชีตจะไม่ได้ Location
const webSheets = new Function(htmlLf.match(/var LINE_SHEET_OPTIONS = \{[\s\S]*?\n    \};/)[0] + '\nreturn LINE_SHEET_OPTIONS;')();
Object.keys(webSheets).forEach(function(line) {
  assert.deepStrictEqual(P[line].sheets, webSheets[line], 'ชีตของไลน์ ' + line + ' ต้องตรงกับ LINE_SHEET_OPTIONS');
});

// ── หมวดพิมพ์ไม่ตรงกันก็ต้องจับได้ ───────────────────────────────────────────────
assert.strictEqual(be.locationPlanCategoryKey(' PNEUMATIC '), 'pneumatic');
assert.strictEqual(be.locationPlanCategoryKey('Electric'), 'electrical');
assert.strictEqual(be.locationPlanCategoryKey('Machanical'), 'mechanical');
assert.strictEqual(be.locationPlanCategoryKey('Sensor / Vision'), '');
assert.strictEqual(be.locationPlanCategoryKey(''), '');

// ── แผนการเขียน ─────────────────────────────────────────────────────────────────
const rows = [
  { no: '1', name: 'Cylinder', category: 'Pneumatic', location: '-' },
  { no: '2', name: 'Relay', category: 'Electrical', location: 'sp-h9-a4-02' },
  { no: '3', name: 'Bearing', category: 'Mechanical', location: 'OLD-1' },
  { no: '4', name: 'Glove', category: 'Consumable', location: '' },
  { no: '5', name: 'Belt', category: 'mechanical', location: '' }
];
const all = be.planLocationsForRows(rows, P['H9'].locations, false);
assert.deepStrictEqual(all.assign.map(function(a) { return [a.index, a.to, a.from]; }),
  [[0, 'SP-H9-A4-01', ''], [2, 'SP-H9-A4-03', 'OLD-1'], [4, 'SP-H9-A4-03', '']]);
assert.strictEqual(all.same, 1, 'Location ตรงอยู่แล้ว (ต่างแค่ตัวพิมพ์) ไม่ต้องเขียนซ้ำ');
assert.deepStrictEqual(all.skipped.map(function(s) { return s.no; }), ['4'], 'หมวดอื่นต้องข้าม ไม่เดา Location ให้');
const keep = be.planLocationsForRows(rows, P['H9'].locations, true);
assert.deepStrictEqual(keep.assign.map(function(a) { return a.no; }), ['1', '5']);
assert.strictEqual(keep.kept, 1, 'ติ๊ก "ไม่ทับ" แล้วรายการที่มี Location ต้องคงไว้');

// ── backend: Admin เท่านั้น, เขียนชีตละครั้ง, ลงทะเบียนการแก้ไข ──────────────────────
const fn = grabTop(backend, 'bulkAssignLocationsByCategory');
assert(fn.includes('requireAdminUser(payload)'));
assert(fn.includes('if (dryRun || !result.assign.length) return;'), 'ดูตัวอย่างต้องไม่เขียนชีต');
assert(/getRange\(ctx\.headerRowIndex \+ 2, locCol \+ 1, colValues\.length, 1\)\.setValues\(colValues\)/.test(fn));
assert(fn.includes('invalidateSparePartsLiteCache(sheetName)'));
assert(fn.includes('getItemAuditSheet()'));
assert(backend.includes("if (action === 'bulkAssignLocationsByCategory') return respond(bulkAssignLocationsByCategory(body), e);"));

// ── หน้าเว็บ: แผงใน Admin > ระบบ + เรียงหน้า Stock ตาม Location ────────────────────
assert(htmlLf.includes('id="locPlanPreviewBtn"'));
assert(htmlLf.includes('id="locPlanApplyBtn"'));
assert(htmlLf.includes("action: 'bulkAssignLocationsByCategory', dryRun: dryRun"));
assert.strictEqual((htmlLf.match(/<option value="location_asc">📍 Location<\/option>/g) || []).length, 2, 'ตัวเลือกเรียงต้องมีทั้งจอมือถือและจอใหญ่');
assert(htmlLf.includes("if (selectedSort === 'location_asc') filtered.sort(function(a,b){ return compareItemLocation(a, b)"));

console.log('location-plan-by-category: ok');

// Tailwind 2.2 (CDN) ไม่มีสี violet — ปุ่มยืนยันเคยเป็นตัวขาวบนพื้นขาว มองไม่เห็น
const locPanel = htmlLf.slice(htmlLf.indexOf('<!-- จัด Location ตามหมวด'), htmlLf.indexOf('id="locPlanMessage"'));
assert(!/class="[^"]*violet-/.test(locPanel), 'แผงจัด Location ห้ามใช้คลาสสี violet');
assert(/id="locPlanApplyBtn"[^>]*bg-indigo-600/.test(locPanel), 'ปุ่มยืนยันต้องมีพื้นสีที่ Tailwind 2 มีจริง');

// ── ฟอร์มเพิ่ม/แก้ไข: Location เป็น Dropdown ตามแผน ไลน์ × หมวด ─────────────────────────
const webPlanSrc = htmlLf.match(/^ {4}var LOCATION_PLAN_BY_LINE = \{[\s\S]*?\n {4}\};/m);
assert(webPlanSrc, 'หน้าเว็บต้องมี LOCATION_PLAN_BY_LINE');
const ui = new Function([
  webPlanSrc[0],
  grab(htmlLf, 'escHtml'),
  grab(htmlLf, 'normalizeLocationValue'),
  grab(htmlLf, 'locationPlanCategoryKey'),
  grab(htmlLf, 'suggestedLocationFor'),
  "var LOC_CATEGORY_LABELS = { pneumatic: 'Pneumatic', electrical: 'Electrical', mechanical: 'Mechanical' };",
  grab(htmlLf, 'buildLocationOptions'),
  'return { LOCATION_PLAN_BY_LINE, suggestedLocationFor, buildLocationOptions };'
].join('\n'))();
Object.keys(P).forEach(function(line) {
  assert.deepStrictEqual(ui.LOCATION_PLAN_BY_LINE[line], P[line].locations, 'แผน Location หน้าเว็บกับ Backend ต้องตรงกัน: ' + line);
});
assert.strictEqual(ui.suggestedLocationFor('Coil Winding', 'Machanical'), 'SP-CL-A1-01');
assert.strictEqual(ui.suggestedLocationFor('Arc Chute', 'General'), '');
// ยังไม่มี Location → เลือกตามหมวดให้ · ไลน์ของรายการอยู่กลุ่มแรก
let o = ui.buildLocationOptions('Arc Chute', 'Pneumatic', '-');
assert.strictEqual(o.value, 'SP-AC-A3-02');
assert.strictEqual(o.groups[0].label, 'Arc Chute (ไลน์นี้)');
assert.deepStrictEqual(o.groups[0].options.map(function(x) { return x.value; }), ['SP-AC-A3-01', 'SP-AC-A3-02', 'SP-AC-A3-03']);
assert.strictEqual(o.groups.length, 4, 'ไลน์อื่นยังเลือกได้');
// มี Location อยู่แล้ว → คงค่าเดิม แม้ไม่อยู่ในแผน (ไม่ทิ้งข้อมูล)
o = ui.buildLocationOptions('H9', 'Electrical', 'old-1');
assert.strictEqual(o.value, 'OLD-1');
assert.deepStrictEqual(o.extra, [{ value: 'OLD-1', label: 'OLD-1 (เดิม)' }]);
o = ui.buildLocationOptions('H9', 'General', '');
assert.strictEqual(o.value, '', 'หมวดอื่นไม่เดาให้');
assert(htmlLf.includes("fillLocationSelect(document.getElementById('eLocation'), lineOfItem(item), item.category, item.location);"));
assert(/\['mCategory', 'mLine', 'mSheet'\]\.forEach[\s\S]{0,120}addEventListener\('change', syncManage\)/.test(htmlLf));
