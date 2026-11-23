// หน้า PR Report ขั้นตอนที่ 3: แก้ไขรายการในการ์ดได้ (ชื่อ / ยี่ห้อ / รุ่น / หน่วย / ประเภท / ไลน์)
// ค่าที่แก้ต้องไหลไปที่ createPR (part_name/brand/model/unit/category) — แก้เฉพาะใน PR ไม่แตะสต็อก
const { html, grabFn, assert } = require('./_export-extract');

const esc = function(str) { return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); };
const cats = html.match(/var PR_EDIT_CATEGORIES = [^\n]+/)[0];
const edit = new Function('escHtml', cats + '\n' + grabFn('prCandidateEditHtml') + '\nreturn prCandidateEditHtml;')(esc);

const out = edit({ name: 'Run-Down Adapter "RDA"', brand: '', model: '', unit: '', category: 'Machanical', line: 'H9', priority: 'Medium' }, 3);
['data-pr-name="3"', 'data-pr-brand="3"', 'data-pr-model="3"', 'data-pr-unit="3"', 'data-pr-category="3"', 'data-pr-line="3"',
  'data-pr-machine="3"', 'data-pr-purpose="3"', 'data-pr-priority="3"', 'data-pr-edit="3"', '✏️ แก้ไขรายการ'].forEach(function(t) {
  assert(out.includes(t), 'ฟอร์มแก้ไขต้องมี: ' + t);
});
assert(out.includes('value="Run-Down Adapter &quot;RDA&quot;"'), 'ชื่อที่มี " ต้อง escape ใน value');
assert(out.includes('value="PCS"'), 'หน่วยว่าง → PCS');
assert(/<option value="Machanical" selected>/.test(out), 'หมวดเดิมจากสต็อกต้องเลือกค้างไว้');
assert(/<option value="H9" selected>/.test(out));
assert(!/ open>/.test(out.split('<summary')[0]), 'ปิดไว้เป็นค่าเริ่มต้น');
assert(/ open>/.test(edit({ name: 'x', _editOpen: true }, 0).split('<summary')[0]), 'เปิดค้างไว้หลัง re-render');
const tools = edit({ name: 'x', category: 'Tools' }, 0);
assert(/<option value="Tools" selected>/.test(tools) && !tools.includes('หมวด Tools'), 'ประเภทมาตรฐานไม่ซ้ำ');

// การ์ดใช้ฟอร์มนี้ และหัวการ์ด escape ชื่อ + มีจุดให้ patch สด
assert(html.includes('+ prCandidateEditHtml(c, idx)'));
assert(html.includes("data-pr-name-text=\"'+idx+'\" class=\"font-semibold text-slate-800\">'+escHtml(c.name||'-')+'</p>'"));
// handler รับฟิลด์ใหม่
['c.name = e.target.value;', 'c.brand = e.target.value;', 'c.model = e.target.value;', "c.unit = String(e.target.value || '').trim() || 'PCS';",
  'c.category = e.target.value; renderPrCandidates();', 'd.prName || d.prBrand || d.prModel || d.prUnit || d.prCategory'].forEach(function(t) {
  assert(html.includes(t), 'input handler ต้องมี: ' + t);
});
// createPR ส่งค่าที่แก้แล้ว
['part_name: item.name ||', 'brand: item.brand ||', 'model: item.model ||', 'category: item.category ||', "unit: item.unit || 'PCS'"].forEach(function(t) {
  assert(html.includes(t), 'createPR ต้องส่ง: ' + t);
});

console.log('pr-candidate-edit: OK');
