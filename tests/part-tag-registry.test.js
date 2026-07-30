const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ── Backend: config + sheet wiring ────────────────────────────────────────────
assert(backend.includes("SPARE_APP_CONFIG.partTagsSheetName"));
assert(backend.includes("SPARE_APP_CONFIG.partTagConfigSheetName"));
assert(backend.includes('var PART_TAG_HEADERS ='));
assert(backend.includes("var PART_TAG_CONFIG_HEADERS = ['Category', 'Require Tag', 'Prefix', 'Start Number', 'Digits', 'Updated By', 'Updated At']"));
// สถานะต้องใช้คำเดียวกับ CWM System ที่หน้างานใช้อยู่ (ติดตั้ง → ถอด)
assert(backend.includes("var PART_TAG_STATUSES = ['รอติดตั้ง', 'ติดตั้งแล้ว', 'ถอดแล้ว', 'ชำรุด/ทิ้ง', 'คืนคลัง']"));
assert(backend.includes("'Installed At', 'Removed At'"), 'ต้องเก็บวันติดตั้ง/วันถอด');
// ติดตั้งใหม่ต้องล้างวันถอดเพื่อเริ่มนับรอบใหม่
assert(backend.includes('installedAt = now;') && backend.includes("removedAt = '';"));

// ออกเลขเฉพาะขาเบิกออก (signedQty < 0) — รับเข้าต้องไม่ออกเลข
assert(backend.includes('if (signedQty < 0) {'));
assert(backend.includes('partTagResult = createPartTagsForIssue(payload, issuedBy, Math.abs(signedQty));'));
assert(backend.includes('part_tags: partTagResult.tags'));
assert(backend.includes('part_tag_warning: partTagResult.skipped_reason'));

// เลขเริ่มต้น/จำนวนหลักต่อ category ต้องถูกส่งต่อไปจนถึงจุดออกเลขจริง —
// นี่คือกลไก "สานต่อ" เลขจากระบบเดิม (เช่น CWM System ใช้ SC-001..SC-039 นอกระบบนี้)
assert(backend.includes('nextPartTagRunning(existing, rule.prefix, rule.start_number)'));
assert(backend.includes('padPartTagRunning(running + i, rule.digits)'));

// action ต้องถูก dispatch ครบทั้ง 4 ตัว
['getPartTagConfig', 'setPartTagConfig', 'getPartTags', 'updatePartTagStatus'].forEach(function (action) {
  assert(backend.includes("action === '" + action + "'"), 'missing dispatch for ' + action);
});

// ── Backend: logic ของฟังก์ชันล้วน (โหลดมารันจริง) ─────────────────────────────
const sandbox = {};
const backendLf = backend.replace(/\r\n/g, '\n');
const pureFns = ['defaultPartTagPrefix', 'nextPartTagRunning', 'padPartTagRunning', 'normalizePartTagCategory', 'rowToPartTagConfig', 'partTagLineFromProcess'];
pureFns.forEach(function (name) {
  const re = new RegExp('^function[ ]+' + name + '\\([\\s\\S]*?\\n}$', 'm');
  const match = backendLf.match(re);
  assert(match, 'cannot extract ' + name);
  sandbox[name] = new Function('defaultPartTagPrefix', match[0] + '\nreturn ' + name + ';')(sandbox.defaultPartTagPrefix);
});

assert.strictEqual(sandbox.defaultPartTagPrefix('Mechanical'), 'MECH');
assert.strictEqual(sandbox.defaultPartTagPrefix('Coil Winding'), 'COIL');
assert.strictEqual(sandbox.defaultPartTagPrefix('อะไหล่'), 'TAG', 'ชื่อไทยล้วนต้อง fallback เป็น TAG');
assert.strictEqual(sandbox.defaultPartTagPrefix(''), 'TAG');

assert.strictEqual(sandbox.padPartTagRunning(1), '00001', 'ไม่ระบุ digits ต้อง fallback เป็น 5 หลัก');
assert.strictEqual(sandbox.padPartTagRunning(123), '00123');
assert.strictEqual(sandbox.padPartTagRunning(123456), '123456', 'เลขเกิน 5 หลักต้องไม่ถูกตัด');
assert.strictEqual(sandbox.padPartTagRunning(40, 3), '040', 'digits=3 ต้องได้ 3 หลักตามระบบเดิม (เช่น SC-040)');
assert.strictEqual(sandbox.padPartTagRunning(4000, 3), '4000', 'เลขเกิน digits ที่กำหนดต้องไม่ถูกตัด');

// เลขวิ่งต้องต่อจากค่ามากสุดของ prefix นั้น และไม่สนใจ prefix อื่น
const rows = [
  ['Tag No', 'Part No'],
  ['MECH-00001', 'A'],
  ['MECH-00007', 'B'],
  ['PIN-00099', 'C'],
  ['MECH-00003', 'D']
];
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'MECH'), 8);
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'PIN'), 100);
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'NEW'), 1, 'prefix ที่ยังไม่มีต้องเริ่มที่ 1 เมื่อไม่ตั้งเลขเริ่มต้น');
assert.strictEqual(sandbox.nextPartTagRunning([['Tag No']], 'MECH'), 1, 'ชีทเปล่าต้องเริ่มที่ 1 เมื่อไม่ตั้งเลขเริ่มต้น');
// prefix ที่เป็น substring กันต้องไม่ปนกัน (MECH- ไม่ใช่ ME-)
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'ME'), 1);

// "สานต่อ" เลขจากระบบเดิม: prefix ที่ยังไม่เคยออกในทะเบียนนี้เลย ต้องเริ่มจาก startNumber ที่ตั้งไว้
assert.strictEqual(sandbox.nextPartTagRunning([['Tag No']], 'SC', 40), 40, 'ชีทเปล่า + ตั้งเลขเริ่มต้น 40 ต้องเริ่มที่ 40 ไม่ใช่ 1');
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'SC', 40), 40, 'prefix ใหม่ที่ยังไม่มีในทะเบียนต้องใช้เลขเริ่มต้นแม้ prefix อื่นมีข้อมูลอยู่แล้ว');
// ถ้าเคยออกเลขของ prefix นี้ไปแล้ว (max>0) ต้องต่อจาก max เสมอ ไม่ใช่วนกลับไปที่ startNumber
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'MECH', 999), 8, 'มีเลขอยู่แล้วต้องต่อจาก max ไม่สนใจ startNumber');

assert.strictEqual(sandbox.normalizePartTagCategory('  Mechanical '), 'mechanical');

// คอลัมน์ Line ต้องไม่มีโน๊ตของการเบิกติดมา
assert.strictEqual(sandbox.partTagLineFromProcess('Coil Winding | note: เปลี่ยนตามรอบ'), 'Coil Winding');
assert.strictEqual(sandbox.partTagLineFromProcess('H9'), 'H9');
assert.strictEqual(sandbox.partTagLineFromProcess(''), '');

const cfg = sandbox.rowToPartTagConfig(['Mechanical', 'TRUE', '', '', '', 'admin', '2026-07-30']);
assert.strictEqual(cfg.require_tag, true);
assert.strictEqual(cfg.prefix, 'MECH', 'prefix ว่างต้อง fallback จากชื่อ category');
assert.strictEqual(cfg.start_number, 1, 'ไม่ระบุเลขเริ่มต้นต้อง fallback เป็น 1');
assert.strictEqual(cfg.digits, 5, 'ไม่ระบุจำนวนหลักต้อง fallback เป็น 5');

const cfgSeeded = sandbox.rowToPartTagConfig(['Stripper Cutter', 'TRUE', 'SC', '40', '3', 'admin', '2026-07-30']);
assert.strictEqual(cfgSeeded.prefix, 'SC');
assert.strictEqual(cfgSeeded.start_number, 40, 'ต้องอ่านเลขเริ่มต้นที่ admin ตั้งไว้ (สานต่อจาก SC-039 เดิม)');
assert.strictEqual(cfgSeeded.digits, 3, 'ต้องอ่านจำนวนหลักที่ admin ตั้งไว้ (SC-040 = 3 หลัก)');

assert.strictEqual(sandbox.rowToPartTagConfig(['Mechanical', 'FALSE', 'MT', '1', '5', 'a', 'b']).require_tag, false);
assert.strictEqual(sandbox.rowToPartTagConfig(['Mechanical', '1', 'MT', '1', '5', 'a', 'b']).require_tag, true);

// ── Frontend: UI + การรับเลขกลับมา ────────────────────────────────────────────
assert(html.includes('id="partTagIssuedModal"'));
assert(html.includes('id="partTagRegistryModal"'));
assert(html.includes('id="partTagConfigList"'));
assert(html.includes('#partTagIssuedModal { z-index: 9800; }'), 'สรุปเลขต้องอยู่เหนือ Issue Cart');
assert(html.includes('if (res.part_tags && res.part_tags.length) {'));
assert(html.includes('showIssuedPartTagsSummary(issuedTagGroups);'));
// ต้อง degrade เงียบๆ ถ้า backend ยังไม่ deploy
assert(html.includes('partTagConfigState.supported = false;'));
assert(html.includes('if (!partTagConfigState.supported) return null;'));

// Admin: ช่องตั้งเลขเริ่มต้น (สานต่อจากระบบเดิม เช่น CWM System) ต้องแปลง "040" เป็น start=40, digits=3
assert(html.includes('id="partTagConfigStartInput"'));
assert(html.includes("var startRaw = String((startEl && startEl.value) || '').replace(/\\D/g, '');"));
assert(html.includes("var startNumber = startRaw ? parseInt(startRaw, 10) : 1;"));
assert(html.includes("var digits = startRaw ? startRaw.length : 5;"));
// ปุ่มเปิด/ปิดแท็กเร็วๆ ต้องไม่รีเซ็ต prefix/เลขเริ่มต้น/จำนวนหลักที่เคยตั้งไว้
assert(html.includes('current ? current.startNumber : 1, current ? current.digits : 5'));

console.log('Part tag registry checks passed');
