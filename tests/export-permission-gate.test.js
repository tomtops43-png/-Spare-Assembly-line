// สิทธิ์และร่องรอยของ Export Center
// ไฟล์ที่ออกจากหน้านี้ออกไปอยู่นอกองค์กร (มือลูกค้า/ผู้จัดการ) จึงต้อง
//  1) กันคนที่ไม่มีสิทธิ์ ทั้งฝั่ง UI และฝั่งเซิร์ฟเวอร์ (ห้ามพึ่ง UI อย่างเดียว)
//  2) คนที่ถูกจำกัดไลน์ ต้องดึงได้แค่ไลน์ตัวเอง — บังคับที่ server
//  3) บันทึกทุกครั้งที่ดาวน์โหลดว่าใครดึงอะไรไป
const { html, backend, grabFn, assert } = require('./_export-extract');

// ── ฝั่งเซิร์ฟเวอร์: ทุก action ของ export ต้องผ่านการตรวจสิทธิ์ ────────
const exportActions = ['exportManifest', 'exportPrBundle', 'exportAuditTrails', 'exportRawSheets', 'logExportEvent'];
exportActions.forEach(function(name) {
  const at = backend.indexOf('function ' + name + '(');
  assert(at > -1, 'Backend ต้องมีฟังก์ชัน ' + name);
  const body = backend.slice(at, at + 1400);
  assert(/requireExportAccess|requireAdminUser/.test(body), name + ' ต้องตรวจสิทธิ์ก่อนคืนข้อมูล');
});
// exportUserRoster เข้มกว่า: Admin เท่านั้น
const rosterAt = backend.indexOf('function exportUserRoster(');
assert(/requireAdminUser/.test(backend.slice(rosterAt, rosterAt + 400)), 'exportUserRoster ต้องเป็น Admin เท่านั้น');

// ── requireExportAccess ต้องยอมรับ 3 ทาง ──────────────────────────────
// export_data เป็นเกณฑ์หลัก แต่ getRoleDefaultPermissions() ไม่เคยมีคีย์นี้มาก่อน
// ผู้ใช้ที่สร้างไว้ก่อนฟีเจอร์นี้จะไม่มีใน permissions_json ถ้าเช็คเข้มตัวเดียว Admin ปัจจุบันใช้ไม่ได้
const accessSrc = grabFn('requireExportAccess', backend, 0);
assert(/hasPermissionForUser\(user, 'export_data'\)/.test(accessSrc), 'ต้องรับ export_data');
assert(/hasPermissionForUser\(user, 'view_logs'\)/.test(accessSrc), 'ต้องรับ view_logs (ผู้ใช้เก่าไม่มีคีย์ export_data)');
assert(/normalizeRole\(user\.role\) === 'admin'/.test(accessSrc), 'Admin ต้องผ่านเสมอ');
assert(/throw new Error/.test(accessSrc), 'คนที่ไม่เข้าเกณฑ์ต้องถูกปฏิเสธ');

// role default ใหม่ต้องมี export_data ให้ admin/leader — ผู้ใช้ที่สร้างต่อจากนี้จะได้ติดมาเอง
const roleDefaults = backend.slice(backend.indexOf('function getRoleDefaultPermissions('), backend.indexOf('function parsePermissions('));
assert((roleDefaults.match(/export_data: true/g) || []).length >= 2, 'admin และ leader ควรได้ export_data ติดมาโดยค่าเริ่มต้น');

// ── คนที่ถูกจำกัดไลน์ ต้องดึงได้แค่ไลน์ตัวเอง (บังคับที่ server) ────────
const allowedLineSrc = grabFn('exportAllowedLine', backend, 0);
assert(/user\.line/.test(allowedLineSrc), 'ต้องอ่านไลน์ที่ผูกกับผู้ใช้');
const prSrc = grabFn('exportPrBundle', backend, 0);
assert(/var forcedLine = exportAllowedLine\(user\);/.test(prSrc), 'exportPrBundle ต้องบังคับไลน์จากสิทธิ์ผู้ใช้');
assert(/if \(forcedLine\) line = forcedLine;/.test(prSrc), 'ไลน์ที่ผู้ใช้ส่งมาต้องถูกทับด้วยไลน์ที่มีสิทธิ์ ไม่ใช่เชื่อค่าจากฝั่งเว็บ');
const rawSrc = grabFn('exportRawSheets', backend, 0);
assert(/exportAllowedLine\(user\)/.test(rawSrc), 'exportRawSheets ต้องบังคับไลน์ด้วย');
// ชีตดิบต้องอ่านได้เฉพาะที่อยู่ใน whitelist — ไม่ใช่ให้ระบุชื่อชีตอิสระ
assert(/EXPORT_RAW_SHEET_WHITELIST\.indexOf\(name\) === -1/.test(rawSrc), 'exportRawSheets ต้องปฏิเสธชีตที่ไม่อยู่ใน whitelist');
assert(/denied: true/.test(rawSrc), 'ชีตที่ไม่อนุญาตต้องตอบว่าถูกปฏิเสธ ไม่ใช่เงียบ');

// ── ทุก action ต้องมี route ทั้ง doGet และ doPost ───────────────────────
// ฝั่งเว็บใช้ JSONP (GET) เป็นหลักเพราะ API อยู่คนละ origin แต่ requestApi() ลอง POST ก่อน
// ถ้าลงทะเบียนแค่ทางเดียวจะพังเป็นบางเครื่อง/บางเน็ตแบบหาสาเหตุยาก
exportActions.concat(['exportUserRoster']).forEach(function(name) {
  const get = new RegExp("if \\(action === '" + name + "'\\) return respond\\(" + name + "\\(e\\.parameter\\), e\\);");
  const post = new RegExp("if \\(action === '" + name + "'\\) return respond\\(" + name + "\\(body\\), e\\);");
  assert(get.test(backend), 'ขาด route ใน doGet: ' + name);
  assert(post.test(backend), 'ขาด route ใน doPost: ' + name);
});

// ── ฝั่งเว็บ: เกณฑ์สิทธิ์ต้องตรงกับฝั่งเซิร์ฟเวอร์ ────────────────────
const canExportSrc = grabFn('xpCanExport');
assert(/hasPermission\('export_data'\)/.test(canExportSrc));
assert(/hasPermission\('view_logs'\)/.test(canExportSrc));
assert(/isAdminUser\(\)/.test(canExportSrc));

// แท็บต้องถูกซ่อนถ้าไม่มีสิทธิ์ และเข้าหน้าตรงๆ ก็ต้องถูกเด้งกลับ
assert(/tabExportEl\.classList\.toggle\('hidden', !\(typeof xpCanExport === 'function' && xpCanExport\(\)\)\)/.test(html.replace(/\r\n/g, '\n')), 'setTabStyles ต้องซ่อนแท็บ Export เมื่อไม่มีสิทธิ์');
const switchSrc = grabFn('switchToExportPage');
assert(/if \(!xpCanExport\(\)\)/.test(switchSrc), 'switchToExportPage ต้องเช็คสิทธิ์ก่อนเปิดหน้า');
assert(/return switchToStockPage\(\)/.test(switchSrc), 'ไม่มีสิทธิ์ต้องเด้งกลับหน้า Stock');
assert(/xpAuditCard[\s\S]*isAdminUser\(\)/.test(switchSrc), 'ทะเบียนการส่งออกต้องโผล่เฉพาะ Admin');

// ── ทุกทางที่ดาวน์โหลดไฟล์ ต้องขึ้นทะเบียน ────────────────────────────
['xpRunWorkbook', 'xpRunSelectedExport', 'xpRunPdfReport', 'xpDownloadReport'].forEach(function(fn) {
  const src = grabFn(fn);
  assert(/xpLogExport\(/.test(src), fn + ' ต้องเรียก xpLogExport เพื่อขึ้นทะเบียนว่าใครดึงอะไรไป');
});
const logExportSrc = grabFn('xpLogExport');
assert(/customerMode/.test(logExportSrc), 'ทะเบียนต้องบันทึกว่าไฟล์นั้นเป็นโหมดลูกค้าหรือไม่');
assert(/from: xpFilters\.from, to: xpFilters\.to/.test(logExportSrc), 'ทะเบียนต้องบันทึกช่วงข้อมูลที่ดึงไป');
assert(/\.catch\(/.test(logExportSrc), 'ถ้าบันทึกทะเบียนไม่ได้ ต้องไม่ทำให้ไฟล์ที่ผู้ใช้ได้ไปแล้วพัง');

// ฝั่งเซิร์ฟเวอร์: logExportEvent ต้องไม่ throw ออกไปข้างนอก
const logEventSrc = grabFn('logExportEvent', backend, 0);
assert(/try \{/.test(logEventSrc) && /catch \(err\)/.test(logEventSrc), 'logExportEvent ต้องกลืน error เอง');
assert(/return \{ status: 'success', logged: false \}/.test(logEventSrc), 'บันทึกไม่ได้ต้องตอบว่าไม่ได้บันทึก แต่ไม่ล้ม request');

// ── ทะเบียนการส่งออกต้องอ่านได้เฉพาะ Admin ────────────────────────────
const auditSrc = grabFn('exportAuditTrails', backend, 0);
assert(/wanted\('exportLog'\) && normalizeRole\(user\.role\) === 'admin'/.test(auditSrc), 'ทะเบียนการส่งออกต้องอ่านได้เฉพาะ Admin');

console.log('export-permission-gate: OK');
