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

// ---- เบิกด่วน (quick issue): เติมเลขลงช่องอัตโนมัติ ไม่ต้องกดปุ่ม ----
assert(html.includes("var seq = formatTagNoSequence(rule, q);"), 'quick issue computes the run from the current qty');
assert(html.includes("input.placeholder = seq ? ('เช่น ' + seq) : 'เลขประจำชิ้น';"), 'placeholder shows the full run');
assert(/if \(seq && \(!typed \|\| typed === quickIssueAutoTagNos\)\) \{\s*input\.value = seq;\s*quickIssueAutoTagNos = seq;/.test(html), 'quick issue auto-fills the field without a button press');
assert(html.includes("if (resetValue) { input.value = ''; quickIssueAutoTagNos = ''; }"), 'opening the modal for another item clears the auto-fill memory');
assert(html.includes('ระบบใส่เลขถัดไปให้แล้ว'), 'hint tells the tech the numbers were filled in for them');
// ปุ่มยังอยู่ แต่เปลี่ยนหน้าที่เป็น "ดึงเลขถัดไปกลับมา" — ทับได้ ไม่บล็อกเหมือนเดิม
assert(!html.includes('มีเลขกรอกไว้แล้ว — ลบออกก่อนถ้าต้องการใช้เลขใหม่'), 'the button no longer refuses when the field is filled');
assert(/var qty = Number\(\(document\.getElementById\('quickIssueQty'\) \|\| \{\}\)\.value \|\| 1\);\s*var seq = formatTagNoSequence\(rule, qty\);\s*input\.value = seq;/.test(html), 'quick issue button fills the whole run');

// ---- Issue Cart: บรรทัดละหลายชิ้นก็ต้องเติมให้ครบอัตโนมัติ ----
assert(html.includes('var tagSeq = formatTagNoSequence(tagRule, qty);'), 'cart row computes the run from its qty');
assert(/if \(tagSeq && \(!savedTagNos\.trim\(\) \|\| savedTagNos === issueCartAutoTagNos\[entry\.key\]\)\) \{/.test(html), 'cart row auto-fills, and only overwrites its own previous auto-fill');
assert(html.includes('issueCartAutoTagNos[entry.key] = tagSeq;'), 'cart remembers what it auto-filled');
assert(html.includes('var seq = formatTagNoSequence(rule, Math.min(entry.qty, safeNum(entry.item.stock)));'), 'cart button fills the run capped at stock');
assert(html.includes('issueCartTagNos[entry.key] = seq;'), 'cart state keeps the full run');

// ---- ห้ามทับเลขที่ช่างพิมพ์เอง + ล้าง state ให้ครบ ----
assert(html.includes('var issueCartAutoTagNos = {};'), 'auto-fill memory declared');
assert(html.includes("var quickIssueAutoTagNos = '';"), 'quick issue auto-fill memory declared');
assert(html.includes('delete issueCartTagNos[key]; delete issueCartAutoTagNos[key];'), 'submitted rows clear both maps');
assert(/function removeFromIssueCart\(key\) \{[\s\S]{0,200}delete issueCartAutoTagNos\[key\];/.test(html), 'removing a cart row clears its auto-fill memory');

// ---- เลขที่ใช้ไปแล้วต้องไม่ถูกเติมซ้ำรอบหน้า ----
assert(/showIssuedPartTagsSummary\(issuedTagGroups\);[\s\S]{0,200}loadPartTagConfig\(true\);/.test(html), 'issue cart refreshes next_tag_no after issuing');
assert(html.includes('if (res.part_tags && res.part_tags.length) loadPartTagConfig(true);'), 'quick issue refreshes next_tag_no after issuing');

// ---- ตัวตรวจตอนบันทึกยังต้องบังคับให้จำนวนเลขตรงกับจำนวนที่เบิก ----
assert(html.includes('if (quickNos.length !== qty) {'), 'submit still validates tag count against qty');

console.log('part-tag-sequence-by-qty: all assertions passed');
