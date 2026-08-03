const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// แก้ไขรายการอะไหล่ของไลน์ Coil Winding: ช่อง "Machine Model / Coil Size" เดิมเป็น
// text input พิมพ์เอง — เปลี่ยนเป็น dropdown ให้เหมือนช่องเดียวกันในฟอร์ม "เพิ่มรายการใหม่" (#mCoilSize)

// ---- HTML: dropdown แทน text input ----
assert(!/id="eCoilSize" class="[^"]*"\s*\/>/.test(html), 'eCoilSize is no longer a bare text input');
// ตัวเลือกต้องลงท้าย A ให้ตรงกับฟิลเตอร์ Coil Size (ดู pr เรื่อง dropdown ไม่ตรงกัน)
assert(/<select id="eCoilSize"[^>]*>[\s\S]{0,400}<option>10A<\/option>[\s\S]{0,100}<option>16A<\/option>[\s\S]{0,100}<option>20A<\/option>[\s\S]{0,100}<option>25\/32A<\/option>[\s\S]{0,100}<option>Common<\/option>[\s\S]{0,100}<option>Other<\/option>/.test(html), 'eCoilSize select has the same preset options as the add-part form (#mCoilSize)');
assert(html.includes('id="eCoilSizeOther"'), 'has a fallback free-text input for custom Coil Size values');

// ---- JS: populate on open ----
const openEditModalBody = (function() {
  const start = html.indexOf('function openEditModal(item)');
  const end = html.indexOf('\n    }', start);
  assert(start > -1 && end > start, 'can slice openEditModal body');
  return html.slice(start, end);
})();
assert(openEditModalBody.includes("coilPresetOptions.indexOf(coilVal) === -1"), 'unmatched legacy values fall back to the Other slot instead of being dropped');
assert(openEditModalBody.includes("eCoilSize.value = 'Other'"), 'legacy free-text values select Other and populate the fallback input');

// ---- JS: submit reads Other input when selected ----
assert(/var sel = document\.getElementById\('eCoilSize'\);[\s\S]{0,200}if \(v === 'Other'\) \{ var other = document\.getElementById\('eCoilSizeOther'\)/.test(html), 'submit payload reads the Other text input when Other is selected');

// ---- JS: change listener toggles the Other input, mirroring #mCoilSize's listener ----
assert(/eCoilSizeSelect\.addEventListener\('change', function\(\)\{ if\(!eCoilSizeOtherInput\) return; eCoilSizeOtherInput\.classList\.toggle\('hidden', eCoilSizeSelect\.value !== 'Other'\); \}\)/.test(html), 'eCoilSize select toggles the Other input visibility on change');

console.log('edit-coil-size-dropdown: all assertions passed');
