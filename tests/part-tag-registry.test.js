const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ── Backend: config + sheet wiring ────────────────────────────────────────────
assert(backend.includes("SPARE_APP_CONFIG.partTagsSheetName"));
assert(backend.includes("SPARE_APP_CONFIG.partTagConfigSheetName"));
assert(backend.includes('var PART_TAG_HEADERS ='));
assert(backend.includes('var PART_TAG_CONFIG_HEADERS ='));
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

assert.strictEqual(sandbox.padPartTagRunning(1), '00001');
assert.strictEqual(sandbox.padPartTagRunning(123), '00123');
assert.strictEqual(sandbox.padPartTagRunning(123456), '123456', 'เลขเกิน 5 หลักต้องไม่ถูกตัด');

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
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'NEW'), 1, 'prefix ที่ยังไม่มีต้องเริ่มที่ 1');
assert.strictEqual(sandbox.nextPartTagRunning([['Tag No']], 'MECH'), 1, 'ชีทเปล่าต้องเริ่มที่ 1');
// prefix ที่เป็น substring กันต้องไม่ปนกัน (MECH- ไม่ใช่ ME-)
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'ME'), 1);

assert.strictEqual(sandbox.normalizePartTagCategory('  Mechanical '), 'mechanical');

// คอลัมน์ Line ต้องไม่มีโน๊ตของการเบิกติดมา
assert.strictEqual(sandbox.partTagLineFromProcess('Coil Winding | note: เปลี่ยนตามรอบ'), 'Coil Winding');
assert.strictEqual(sandbox.partTagLineFromProcess('H9'), 'H9');
assert.strictEqual(sandbox.partTagLineFromProcess(''), '');

const cfg = sandbox.rowToPartTagConfig(['Mechanical', 'TRUE', '', 'admin', '2026-07-30']);
assert.strictEqual(cfg.require_tag, true);
assert.strictEqual(cfg.prefix, 'MECH', 'prefix ว่างต้อง fallback จากชื่อ category');
assert.strictEqual(sandbox.rowToPartTagConfig(['Mechanical', 'FALSE', 'MT', 'a', 'b']).require_tag, false);
assert.strictEqual(sandbox.rowToPartTagConfig(['Mechanical', '1', 'MT', 'a', 'b']).require_tag, true);

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

console.log('Part tag registry checks passed');
