const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// ── ช่อง "ใช้กับเครื่องไหน" ต้องเป็น dropdown จากทะเบียนเครื่องจักร ────────────
// เดิมเป็น <input list="issueCartMachineOptions"> ที่ datalist ถูกเติมจากชื่อเครื่อง
// ที่เคยพิมพ์ไว้ในประวัติ Log → พิมพ์ผิดครั้งเดียวก็กลายเป็นตัวเลือกถาวร และเครื่องที่
// ลงทะเบียนไว้จริงกลับไม่ขึ้นให้เลือกเลย
assert(html.includes('<select id="issueCartMachine"'), 'ต้องเป็น select ไม่ใช่ input');
assert(!html.includes('list="issueCartMachineOptions"'), 'ต้องเลิกใช้ datalist เดิม');
assert(!html.includes('id="issueCartMachineOptions"'), 'ต้องลบ datalist element ทิ้ง');
assert(!html.includes('function updateMachineDatalist'), 'ต้องลบตัวเติม datalist จาก log ทิ้ง');
assert(!/\n\s*updateMachineDatalist\(/.test(html), 'ต้องไม่มีที่เรียก updateMachineDatalist ค้างอยู่');

// ตัวเลือกต้องมาจาก getMachines ของไลน์ปัจจุบัน (ทะเบียนจริง) ไม่ใช่จาก logCache
assert(html.includes('function populateIssueCartMachineOptions()'));
assert(/function populateIssueCartMachineOptions\(\)[\s\S]{0,700}fetchMachinesForLine\(line\)/.test(html),
  'ต้องดึงรายชื่อจาก fetchMachinesForLine ของไลน์ปัจจุบัน');
assert(/function populateIssueCartMachineOptions\(\)[\s\S]{0,400}String\(currentLine \|\| ''\)/.test(html));

// ต้องเติมตัวเลือกตอนเปิด cart ไม่งั้น dropdown จะว่างจนกว่าจะไปหน้าอื่น
assert(/populateIssueCartMachineOptions\(\);[\s\S]{0,80}renderIssueCart\(\);/.test(html),
  'openIssueCart ต้องเรียก populateIssueCartMachineOptions');

// ยังต้องพิมพ์เองได้เผื่อเครื่องที่ยังไม่ได้ลงทะเบียน — ใช้แพทเทิร์นเดียวกับช่อง Reason
assert(html.includes("var ISSUE_CART_MACHINE_OTHER = '__other__';"));
assert(html.includes('id="issueCartMachineOther"'));
assert(html.includes('✏️ อื่นๆ (พิมพ์เอง)'));
assert(html.includes('function syncIssueCartMachineOtherVisibility()'));
assert(html.includes("issueCartMachineSelect.addEventListener('change', syncIssueCartMachineOtherVisibility)"),
  'เปลี่ยนค่าใน select ต้องซ่อน/แสดงช่องพิมพ์เอง');

// ไลน์ที่ยังไม่มีเครื่องลงทะเบียน ต้องบอกให้รู้ ไม่ใช่ปล่อย dropdown ว่างเปล่า
assert(html.includes('ยังไม่มีเครื่องจักรของไลน์นี้ — เพิ่มที่ Admin'));

// ── ค่าที่ส่งไปบันทึกต้องมาจาก helper ตัวเดียว ────────────────────────────────
// ถ้าอ่าน .value ของ select ตรงๆ จะได้ '__other__' ติดไปในประวัติแทนชื่อเครื่องจริง
assert(html.includes('function getIssueCartMachineValue()'));
assert(html.includes('var machineValue = getIssueCartMachineValue();'));
assert(!html.includes("var machineValue = String((machineEl && machineEl.value) || '').trim();"),
  'ต้องไม่อ่าน .value ของ select ตรงๆ แล้ว');
assert(/function getIssueCartMachineValue\(\)[\s\S]{0,400}ISSUE_CART_MACHINE_OTHER[\s\S]{0,200}issueCartMachineOther/.test(html),
  'เลือก "อื่นๆ" ต้องคืนค่าจากช่องพิมพ์เอง ไม่ใช่ค่า sentinel');

// เบิกเสร็จต้องรีเซ็ตช่องเครื่องกลับ ไม่ค้างไปรอบถัดไป
assert(html.includes('function resetIssueCartMachineField()'));
assert(html.includes('resetIssueCartMachineField();'));

console.log('Issue cart machine dropdown checks passed');
