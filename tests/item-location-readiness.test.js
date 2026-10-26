const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(src, name) {
  const m = src.match(new RegExp('^ {0,4}function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n {0,4}\\}', 'm'));
  assert(m, 'ต้องดึง ' + name + ' ออกมาได้');
  return m[0];
}

// ── ทุกรายการจะมี Location: ค้นหา/กรอง/กรอกให้ตรงกัน/ไม่ทำให้ดาวหาย ──────────────
const web = new Function([
  grab(htmlLf, 'getItemIdentityKey'),
  grab(htmlLf, 'normalizeItemIdentityKey'),
  grab(htmlLf, 'hasItemLocation'),
  grab(htmlLf, 'normalizeLocationValue'),
  grab(htmlLf, 'getItemSearchText'),
  'return { getItemIdentityKey, normalizeItemIdentityKey, hasItemLocation, normalizeLocationValue, getItemSearchText };'
].join('\n'))();

const item = { no: '12', subLine: 'Stock for MC', model: '6204ZZ', name: 'Bearing', location: 'SP-H9-A1-1' };
// ย้ายที่เก็บ/เติม Location แล้วคีย์ต้องเท่าเดิม (ดาว ตะกร้าเบิก อะไหล่ของเครื่องจักร ผูกกับคีย์นี้)
assert.strictEqual(web.getItemIdentityKey(item), web.getItemIdentityKey(Object.assign({}, item, { location: 'B2' })));
assert.strictEqual(web.getItemIdentityKey(item), '12::Stock for MC::6204ZZ::Bearing');
// คีย์รุ่นเก่า (มี location ต่อท้าย) ต้องแปลงกลับมาเป็นคีย์ปัจจุบันได้
assert.strictEqual(web.normalizeItemIdentityKey('12::Stock for MC::6204ZZ::Bearing::-'), '12::Stock for MC::6204ZZ::Bearing');
assert.strictEqual(web.normalizeItemIdentityKey('12::Stock for MC::6204ZZ::Bearing'), '12::Stock for MC::6204ZZ::Bearing');
assert(/var target = normalizeItemIdentityKey\(key\);/.test(htmlLf), 'findItemByIdentityKey ต้องรับคีย์รุ่นเก่าได้');
assert(/var newKey = normalizeItemIdentityKey\(oldKey\);[\s\S]{0,300}if \(migrated\) saveStarredToStorage\(\);/.test(htmlLf), 'ดาวใน localStorage ต้องถูกย้ายไปคีย์ใหม่');

// รูปแบบเดียวกัน: a1-1 กับ A1-1 ต้องเป็นช่องเดียวกัน
assert.strictEqual(web.normalizeLocationValue('  sp-h9  a1-1 '), 'SP-H9 A1-1');
assert.strictEqual(web.normalizeLocationValue(null), '');
assert.strictEqual(web.hasItemLocation({ location: '-' }), false);
assert.strictEqual(web.hasItemLocation({ location: ' ' }), false);
assert.strictEqual(web.hasItemLocation({ location: 'A1' }), true);

// ช่องค้นหาหน้า Stock ต้องหาด้วย Location ได้ (แต่ไม่ให้ "-" ทำให้ทุกตัวติดผล)
assert(web.getItemSearchText(Object.assign({}, item)).indexOf('sp-h9-a1-1') > -1);
assert.strictEqual(web.getItemSearchText({ name: 'X', location: '-' }).indexOf('-'), -1);

// ปุ่ม "ยังไม่มี Location" + ตัวกรอง
assert(htmlLf.includes('id="noLocationFilterBtn"'));
assert(htmlLf.includes('if (noLocationOnlyState && hasItemLocation(item)) return false;'));
assert(htmlLf.includes("'ยังไม่มี Location (' + count + ')'"));
// มีรายการ Location เดิมให้เลือก + บันทึกเป็นรูปแบบเดียวกัน
assert(htmlLf.includes('<datalist id="locationSuggestions"></datalist>'));
assert(/\['eLocation', 'mLocation'\]\.forEach[\s\S]{0,200}setAttribute\('list', 'locationSuggestions'\)/.test(htmlLf));
assert(htmlLf.includes("location: normalizeLocationValue(document.getElementById('eLocation').value),"));
assert(htmlLf.includes("location: normalizeLocationValue(document.getElementById('mLocation').value),"));

// ── Backend: ชีตเป็นที่เก็บเดียว ──────────────────────────────────────────────
const be = new Function([grab(backend, 'normalizeLocationText'), grab(backend, 'resolveItemLocation'),
  'return { normalizeLocationText, resolveItemLocation };'].join('\n'))();
assert.strictEqual(be.resolveItemLocation('A1-1', 'OLD-9'), 'A1-1', 'แก้ในชีตแล้วต้องชนะ override รุ่นเก่า');
assert.strictEqual(be.resolveItemLocation('', 'OLD-9'), 'OLD-9', 'ชีตว่างค่อยใช้ override (ข้อมูลเก่า)');
assert.strictEqual(be.resolveItemLocation('-', ''), '-');
assert.strictEqual(be.resolveItemLocation(undefined, undefined), '-');
assert.strictEqual(be.normalizeLocationText(' a1  -1 '), 'A1 -1');
assert(backend.includes('location: resolveItemLocation(rawLocation, locationOverride),'));
assert(!/setLocationOverride\(sheetName, noValue, values\.location\)/.test(backend), 'บันทึกแล้วต้องไม่เขียน override ซ้ำอีก');
assert(backend.includes("if (key === 'location' && String(nextValue || '').trim() === '') continue;"), 'Location ว่าง = ไม่ลบของเดิม');
assert(backend.includes("location: normalizeLocationText(payload.location),"));

// ── ชุดที่ 2: เรียงแบบตัวเลข + ป้าย 📍 บนการ์ด/ตะกร้าเบิก ───────────────────────────
const web2 = new Function([
  grab(htmlLf, 'hasItemLocation'),
  grab(htmlLf, 'compareItemLocation'),
  grab(htmlLf, 'escHtml'),
  grab(htmlLf, 'renderLocationChip'),
  'return { compareItemLocation, renderLocationChip };'
].join('\n'))();
const sorted = [{ location: 'A10' }, { location: '-' }, { location: 'a2' }, { location: 'A1' }, { location: '' }]
  .sort(web2.compareItemLocation).map(function(x) { return x.location; });
assert.deepStrictEqual(sorted.slice(0, 3), ['A1', 'a2', 'A10'], 'A2 ต้องมาก่อน A10');
assert(sorted.slice(3).every(function(l) { return l === '-' || l === ''; }), 'รายการที่ยังไม่มี Location ต้องอยู่ท้าย');
// หน้านับสต็อก: ลำดับตอนเริ่มรอบนับกับใบที่พิมพ์ต้องใช้ตัวเรียงเดียวกัน
assert.strictEqual((htmlLf.match(/return compareItemLocation\(a, b\) \|\| String\(a\.name\|\|''\)\.localeCompare/g) || []).length, 2);

assert.strictEqual(web2.renderLocationChip({ location: '-' }), '');
const chip = web2.renderLocationChip({ location: 'SP-H9-A1-1<b>' });
assert(chip.includes('📍') && chip.includes('SP-H9-A1-1&lt;b&gt;'), 'ป้ายต้องมี 📍 และ escape ข้อความ');
assert(htmlLf.includes(`(locationTxt ? '<div class="mt-2">' + renderLocationChip(item) + '</div>' : '') +`), 'การ์ด Stock ต้องใช้ป้ายแบบใหม่');
assert(htmlLf.includes(`(hasItemLocation(item) ? '<div class="mt-0.5">' + renderLocationChip(item) + '</div>' : '') +`), 'ตะกร้าเบิกต้องโชว์ Location');

console.log('Item location readiness checks passed');
