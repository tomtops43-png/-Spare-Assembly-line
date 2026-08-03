const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// บั๊กจริง: dropdown "Machine Model / Coil Size" ในฟอร์มแก้ไข/เพิ่มรายการ เขียน option มือ
// เป็น 10 / 16 / 20 / 25/32 (ไม่มี A) แต่ฟิลเตอร์หน้า Stock ใช้ 10A / 16A / 20A / 25/32A
// ผลคือเลือกในฟอร์มแล้วค่าไม่ตรงฟิลเตอร์ และแถวที่เป็น '25/32A' อยู่แล้วเด้งไปช่อง "Other"

// ---- แหล่งความจริงเดียว ----
assert(html.includes("var COIL_SIZE_OPTIONS = ['10A', '16A', '20A', '25/32A', 'Common'];"), 'single canonical Coil Size list');
assert(html.includes('function canonicalCoilSize(value)'), 'canonicalizer exists');
assert(html.includes('function coilSizeOptionsHtml(placeholderLabel, includeOther)'), 'shared option-html builder exists');

// ---- ฟิลเตอร์ต้องอ่านจากลิสต์กลาง ไม่ hardcode ซ้ำ ----
assert(html.includes('var fixed = COIL_SIZE_OPTIONS.slice();'), 'filter builds its options from the shared list');

// ---- ทุก select ในฟอร์มต้องเป็นค่าเดียวกับฟิลเตอร์ (ลงท้าย A) ----
['mCoilSize', 'eCoilSize'].forEach(function(id) {
  const start = html.indexOf('<select id="' + id + '"');
  assert(start > -1, id + ' select exists');
  const block = html.slice(start, html.indexOf('</select>', start));
  ['10A', '16A', '20A', '25/32A', 'Common'].forEach(function(v) {
    assert(block.includes('<option>' + v + '</option>'), id + ' offers ' + v);
  });
  ['10', '16', '20', '25/32'].forEach(function(bad) {
    assert(!block.includes('<option>' + bad + '</option>'), id + ' must not offer the no-A variant "' + bad + '"');
  });
});
// ฟอร์มนำเข้าเลิก hardcode แล้ว ใช้ตัวสร้างกลางตัวเดียวกัน
assert(html.includes("<option value=\"-\">-</option>' + coilSizeOptionsHtml('', true) + '"), 'import default coil size uses the shared builder');

// ---- กันไม่ให้ static HTML หลุดไม่ตรงอีก: sync ตอนโหลด ----
assert(html.includes('function syncCoilSizeSelectOptions()'), 'runtime sync helper exists');
assert(/manageForm\.addEventListener\('submit', submitManageItem\);[\s\S]{0,400}syncCoilSizeSelectOptions\(\);/.test(html), 'options are re-synced at init');

// ---- ค่าเก่าที่ไม่มี A ต้องยังใช้งานได้ (ไม่เด้งไป Other / ไม่หลุดฟิลเตอร์) ----
assert(html.includes("var coilVal = canonicalCoilSize(item.coil_size || item.machine_model || '');"), 'edit modal canonicalizes before matching preset options');
assert(html.includes("(canonicalCoilSize(item.coil_size) || '-') !== selectedCoilSize"), 'stock filter matches canonically');
assert(html.includes('var value = canonicalCoilSize(item.coil_size);'), 'filter option list dedupes legacy values');

// ---- บันทึกแล้วต้องเก็บเป็นรูปแบบมาตรฐาน ----
assert(html.includes('return canonicalCoilSize(v);'), 'edit save canonicalizes coil_size');
assert(html.includes("return canonicalCoilSize(v) || '-';"), 'add save canonicalizes coil_size');
assert(html.includes("coil_size: canonicalCoilSize(row[mapping.coil_size] || importState.defaultCoilSize || '-') || '-',"), 'import canonicalizes coil_size');

// ---- canonicalCoilSize ทำงานถูกจริง (รันโค้ดจริงที่ดึงออกมาจากหน้า) ----
const fnStart = html.indexOf('var COIL_SIZE_OPTIONS =');
const fnEnd = html.indexOf('function coilSizeOptionsHtml');
const canon = new Function(html.slice(fnStart, fnEnd) + '; return canonicalCoilSize;')();
assert.strictEqual(canon('25/32'), '25/32A', "'25/32' → '25/32A'");
assert.strictEqual(canon('10'), '10A', "'10' → '10A'");
assert.strictEqual(canon(' 16a '), '16A', 'trims and fixes case');
assert.strictEqual(canon('25/32A'), '25/32A', 'already-canonical stays put');
assert.strictEqual(canon('common'), 'Common', 'Common is case-normalized');
assert.strictEqual(canon('-'), '', "'-' means empty");
assert.strictEqual(canon(''), '', 'empty stays empty');
assert.strictEqual(canon('MF-DRWG-CWM-PF250-K04-25'), 'MF-DRWG-CWM-PF250-K04-25', 'unknown legacy values are preserved verbatim');

console.log('coil-size-options-match-filter: all assertions passed');
