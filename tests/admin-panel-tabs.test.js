const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ── หน้า Admin ต้องแยกเป็นแท็บ ───────────────────────────────────────────────
// เดิมการ์ดทุกใบกองเรียงกันหน้าเดียว ต้องเลื่อนผ่านทะเบียนเครื่องจักรยาวๆ กว่าจะถึง
// รายชื่อผู้ใช้ ซึ่งเป็นงานหลักของหน้านี้
const TABS = ['users', 'machines', 'tags', 'audit', 'system'];
TABS.forEach(function(t) {
  assert(htmlLf.includes('data-admin-tab="' + t + '"'), 'ต้องมีปุ่มแท็บ ' + t);
  assert(htmlLf.includes('data-admin-panel="' + t + '"'), 'ต้องมีพาเนล ' + t);
});
assert(htmlLf.includes('id="adminTabs"'), 'ต้องมีแถบแท็บ');

// ทุกพาเนลต้องเริ่มที่ hidden แล้วให้ switchAdminTab เปิดเฉพาะตัวที่เลือก
// ไม่งั้นตอนโหลดหน้าจะเห็นทุกแท็บซ้อนกันหมดเหมือนเดิม
const panelOpens = htmlLf.match(/<div data-admin-panel="[a-z]+" class="[^"]*"/g) || [];
assert.strictEqual(panelOpens.length, TABS.length, 'จำนวนพาเนลต้องตรงกับจำนวนแท็บ');
panelOpens.forEach(function(tag) {
  assert(/class="hidden/.test(tag), 'พาเนลต้องเริ่มที่ hidden: ' + tag);
});

// ── โหลดข้อมูลเฉพาะแท็บที่เปิด ──────────────────────────────────────────────
// เดิมเข้าหน้า Admin ทีเดียวยิง API 6 ตัวพร้อมกันทั้งที่ส่วนใหญ่ยังไม่ได้ดู
assert(/function loadAdminTabData\(name\) \{[\s\S]{0,200}if \(adminTabLoaded\[name\]\) return;/.test(htmlLf),
  'ต้องโหลดซ้ำแท็บเดิมไม่ได้');
[
  ["users", 'loadUsers()'],
  ["machines", 'loadMachineManageList()'],
  ["tags", 'loadPartTagGroupsList(true)'],
  ["audit", 'loadItemAudit(false)'],
  ["system", 'loadBackupStatus()']
].forEach(function(pair) {
  const re = new RegExp("name === '" + pair[0] + "'\\) \\{[\\s\\S]{0,260}" + pair[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  assert(re.test(htmlLf), 'แท็บ ' + pair[0] + ' ต้องโหลด ' + pair[1]);
});
// เข้าหน้าใหม่ต้องได้ข้อมูลสด ไม่ใช่ค้างของเก่าจากรอบก่อน
assert(/adminTabLoaded = \{\};\n\s+switchAdminTab\(adminCurrentTab\);/.test(htmlLf),
  'เข้าหน้า Admin ใหม่ต้องล้างแคชแล้วโหลดแท็บที่เปิดอยู่');
// ปุ่มของแท็บผู้ใช้ต้องไม่ลอยอยู่ตอนดูแท็บอื่น
assert(/userActions\.classList\.toggle\('hidden', name !== 'users'\)/.test(htmlLf),
  'ปุ่มผู้ใช้ใหม่/รีเฟรช ต้องซ่อนเมื่อไม่ได้อยู่แท็บผู้ใช้');

// ── ทะเบียนเครื่องจักรต้องค้นหาได้ ─────────────────────────────────────────
// บางไลน์มีเครื่องหลายสิบตัว ถ้าไม่มีช่องค้นหาต้องไล่สายตาทีละแถว
assert(htmlLf.includes('id="machineManageSearch"'), 'ต้องมีช่องค้นหาเครื่องจักร');
assert(htmlLf.includes('id="machineManageCount"'), 'ต้องบอกจำนวนเครื่องที่พบ');
assert(/machineManageSearchEl\.addEventListener\('input'/.test(htmlLf), 'พิมพ์แล้วต้องกรองทันที');
// กรองจากรายการที่โหลดไว้แล้ว ไม่ยิง API ใหม่ทุกตัวอักษร
assert(/machineManageCache = list;/.test(htmlLf) && /machineManageCache\.filter/.test(htmlLf),
  'ต้องกรองจากแคชในเครื่อง');

// ── ฟีเจอร์แจ้งเตือน LINE ถูกตัดออกแล้ว (ยังไม่ได้ใช้งานจริง) ──────────────
// เหลือค้างไว้ = โค้ดตายที่ยังต้องอ่าน/ดูแล และ Admin Panel รกโดยไม่จำเป็น
[
  'lineNotifyStatusBadge', 'lineNotifyTestBtn', 'lineNotifyMessage',
  'loadLineNotifyStatus', 'renderLineNotifyStatus', 'แจ้งเตือนเข้า LINE'
].forEach(function(sym) {
  assert(!htmlLf.includes(sym), 'หน้าเว็บต้องไม่เหลือ ' + sym);
});
[
  'getLineConfig', 'isLineConfigured', 'sendLineMessage', 'notifyLine', 'lineFooter',
  'buildDailyLineDigest', 'sendLineTestMessage', 'getLineStatus', 'LINE_CHANNEL_TOKEN', 'LINE_TARGET_ID'
].forEach(function(sym) {
  assert(!backend.includes(sym), 'backend ต้องไม่เหลือ ' + sym);
});
// งานเช้าและ trigger เดิมต้องไม่ถูกลบตามไปด้วย
assert(/function runDailyAutoJobs\(\)/.test(backend), 'งานเช้าต้องยังอยู่');
assert(/results\.autoPr = runAutoPrJob\(\)/.test(backend), 'Auto-PR ต้องยังอยู่');
assert(/results\.anomaly = runIssueAnomalyScan\(\)/.test(backend), 'สแกนการเบิกผิดปกติต้องยังอยู่');
assert(!/results\.line/.test(backend), 'ต้องไม่เหลือการเรียก LINE ในงานเช้า');
assert(!fs.existsSync('tests/line-notification.test.js'), 'เทสต์ของฟีเจอร์ที่ถูกตัดต้องถูกลบด้วย');

console.log('Admin panel tabs + LINE removal checks passed');
