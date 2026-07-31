const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// ── ช่อง "ใช้กับเครื่องไหน" ต้องเป็น dropdown จากทะเบียนเครื่องจักร ────────────
// เดิม Issue Cart เป็น <input list="issueCartMachineOptions"> ที่ datalist ถูกเติมจาก
// ชื่อเครื่องที่เคยพิมพ์ไว้ในประวัติ Log → พิมพ์ผิดครั้งเดียวก็กลายเป็นตัวเลือกถาวร และ
// เครื่องที่ลงทะเบียนไว้จริงกลับไม่ขึ้นให้เลือกเลย ส่วน "เบิกด่วน" ไม่มีช่องนี้เลยแต่แรก
assert(html.includes('<select id="issueCartMachine"'), 'Issue Cart ต้องเป็น select');
assert(html.includes('<select id="quickIssueMachine"'), 'เบิกด่วนต้องมี select เครื่องจักรด้วย');
assert(!html.includes('list="issueCartMachineOptions"'), 'ต้องเลิกใช้ datalist เดิม');
assert(!html.includes('id="issueCartMachineOptions"'), 'ต้องลบ datalist element ทิ้ง');
assert(!html.includes('function updateMachineDatalist'), 'ต้องลบตัวเติม datalist จาก log ทิ้ง');
assert(!/\n\s*updateMachineDatalist\(/.test(html), 'ต้องไม่มีที่เรียก updateMachineDatalist ค้างอยู่');

// ── ตรรกะต้องใช้ร่วมกันทั้งสองฟอร์ม ไม่ก๊อปวาง ────────────────────────────────
assert(html.includes("var MACHINE_SELECT_OTHER = '__other__';"));
assert(html.includes('function populateMachineSelect(selectId, otherId, line)'));
assert(html.includes('function syncMachineSelectOther(selectId, otherId)'));
assert(html.includes('function getMachineSelectValue(selectId, otherId)'));
assert(html.includes('function resetMachineSelect(selectId, otherId)'));
// ตัวเลือกต้องมาจาก getMachines (ทะเบียนจริง) ไม่ใช่จาก logCache
assert(/function populateMachineSelect\(selectId, otherId, line\)[\s\S]{0,700}fetchMachinesForLine\(String\(line \|\| ''\)\.trim\(\)\)/.test(html),
  'ต้องดึงรายชื่อจาก fetchMachinesForLine ตามไลน์ที่ส่งเข้ามา');

// ── Issue Cart wiring ────────────────────────────────────────────────────────
assert(html.includes("return populateMachineSelect('issueCartMachine', 'issueCartMachineOther', currentLine);"));
assert(/populateIssueCartMachineOptions\(\);[\s\S]{0,80}renderIssueCart\(\);/.test(html),
  'openIssueCart ต้องเรียก populateIssueCartMachineOptions');
assert(html.includes('var machineValue = getIssueCartMachineValue();'));
assert(html.includes('resetIssueCartMachineField();'), 'เบิกเสร็จต้องรีเซ็ตช่องเครื่อง');
assert(html.includes("issueCartMachineSelect.addEventListener('change', syncIssueCartMachineOtherVisibility)"));

// ── Quick Issue (เบิกด่วน) wiring ────────────────────────────────────────────
assert(html.includes('id="quickIssueMachineOther"'));
// ต้องผูกกับไลน์ของ "อะไหล่ชิ้นนั้น" ไม่ใช่ไลน์ที่กำลังเปิดดูอยู่ (อาจคนละไลน์)
assert(html.includes('populateQuickIssueMachineOptions(item.line || currentLine);'),
  'เบิกด่วนต้องโหลดเครื่องตามไลน์ของอะไหล่ชิ้นนั้น');
assert(html.includes('resetQuickIssueMachineField();'));
// payload เดิมไม่มี machine เลย — ต้องส่งไปด้วย ไม่งั้นบันทึกแล้วไม่รู้ว่าใส่เครื่องไหน
assert(html.includes('machine: getQuickIssueMachineValue(),'),
  'payload ของเบิกด่วนต้องส่ง machine ไปด้วย');
assert(/quickIssueMachineSelect\.addEventListener\('change'/.test(html));

// ── ยังพิมพ์เองได้เผื่อเครื่องที่ยังไม่ลงทะเบียน ─────────────────────────────
assert((html.match(/✏️ อื่นๆ \(พิมพ์เอง\)/g) || []).length >= 1);
assert((html.match(/🏭 ระบุชื่อเครื่องเอง/g) || []).length === 2, 'ต้องมีช่องพิมพ์เองทั้งสองฟอร์ม');
// ไลน์ที่ยังไม่มีเครื่องลงทะเบียน ต้องบอกให้รู้ ไม่ใช่ปล่อย dropdown ว่างเปล่า
assert(html.includes('ยังไม่มีเครื่องจักรของไลน์นี้ — เพิ่มที่ Admin'));

// ── ค่าที่ส่งไปบันทึกต้องมาจาก helper ตัวเดียว ────────────────────────────────
// ถ้าอ่าน .value ของ select ตรงๆ จะได้ '__other__' ติดไปในประวัติแทนชื่อเครื่องจริง
assert(/function getMachineSelectValue\(selectId, otherId\)[\s\S]{0,400}MACHINE_SELECT_OTHER[\s\S]{0,200}getElementById\(otherId\)/.test(html),
  'เลือก "อื่นๆ" ต้องคืนค่าจากช่องพิมพ์เอง ไม่ใช่ค่า sentinel');
assert(!html.includes("var machineValue = String((machineEl && machineEl.value) || '').trim();"),
  'ต้องไม่อ่าน .value ของ select ตรงๆ แล้ว');

console.log('Issue cart + quick issue machine dropdown checks passed');
