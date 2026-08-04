const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// เดิม: เบิก 3 ชิ้น แต่ปุ่ม "ใช้เลขถัดไป" ใส่ให้เลขเดียว (SC-040) แล้วตอนบันทึกก็ตีกลับว่า
// "ใส่เลขมา 1 เลข แต่เบิก 3 ชิ้น" ช่างต้องไปไล่เลขต่อเอง
// ตอนนี้: ออกเลขต่อเนื่องให้ครบตามจำนวน และโชว์ให้เห็นตั้งแต่ยังไม่กดปุ่ม

// ---- ตัวออกเลขต่อเนื่อง ----
assert(html.includes('function buildTagNoSequence(rule, qty)'), 'sequence builder exists');
assert(html.includes('function formatTagNoSequence(rule, qty)'), 'formatter exists');

// รันโค้ดจริงที่ดึงออกมาจากหน้าเว็บ
const start = html.indexOf('function buildTagNoSequence(rule, qty)');
const end = html.indexOf('function formatTagNoSequence');
assert(start > -1 && end > start, 'can slice the sequence builder');
const buildTagNoSequence = new Function(html.slice(start, end) + '; return buildTagNoSequence;')();

assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'SC-040' }, 3), ['SC-040', 'SC-041', 'SC-042'], 'เบิก 3 ได้ 3 เลขต่อเนื่อง');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'SC-040' }, 1), ['SC-040'], 'เบิก 1 ได้เลขเดียวเหมือนเดิม');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'SC-099' }, 3), ['SC-099', 'SC-100', 'SC-101'], 'ข้ามหลักแล้วต้องไม่ตัดเลข');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'TAG00008' }, 2), ['TAG00008', 'TAG00009'], 'รองรับ prefix ที่ไม่มีขีดคั่น');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: '' }, 3), [], 'ยังไม่รู้เลขถัดไป = ไม่ออกเลขมั่ว');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'ABC' }, 3), ['ABC'], 'เลขไม่ลงท้ายด้วยตัวเลข ไล่ต่อไม่ได้ ให้แค่ตัวเดียว');
assert.deepStrictEqual(buildTagNoSequence({ nextTagNo: 'SC-040' }, 0), ['SC-040'], 'จำนวนเพี้ยนต้องไม่คืน array ว่าง');

// ---- เบิกด่วน (quick issue): ปุ่มใส่ครบตามจำนวน + โชว์ล่วงหน้า ----
assert(/var qty = Number\(\(document\.getElementById\('quickIssueQty'\) \|\| \{\}\)\.value \|\| 1\);\s*var seq = formatTagNoSequence\(rule, qty\);\s*input\.value = seq;/.test(html), 'quick issue fills the whole run, not just one number');
assert(html.includes("var seq = formatTagNoSequence(rule, q);"), 'quick issue hint computes the run from the current qty');
assert(html.includes('ถ้าของยังไม่มีเลข กด "ใช้เลขถัดไป" จะได้ '), 'hint shows which ids the tech will get');
assert(html.includes("input.placeholder = seq ? ('เช่น ' + seq) : 'เลขประจำชิ้น';"), 'placeholder shows the full run');

// ---- Issue Cart: บรรทัดละหลายชิ้นก็ต้องได้ครบ ----
assert(html.includes('var tagSeq = formatTagNoSequence(tagRule, qty);'), 'cart row computes the run from its qty');
assert(html.includes('var seq = formatTagNoSequence(rule, Math.min(entry.qty, safeNum(entry.item.stock)));'), 'cart button fills the run capped at stock');
assert(html.includes('issueCartTagNos[entry.key] = seq;'), 'cart state keeps the full run');

// ---- ตัวตรวจตอนบันทึกยังต้องบังคับให้จำนวนเลขตรงกับจำนวนที่เบิก ----
assert(html.includes('if (quickNos.length !== qty) {'), 'submit still validates tag count against qty');

console.log('part-tag-sequence-by-qty: all assertions passed');
