const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ── Backend: sheet wiring ──────────────────────────────────────────────────
assert(backend.includes("SPARE_APP_CONFIG.partTagsSheetName"));
assert(backend.includes("SPARE_APP_CONFIG.partTagGroupsSheetName"));
assert(backend.includes("SPARE_APP_CONFIG.partTagGroupItemsSheetName"));
assert(backend.includes('var PART_TAG_HEADERS ='));
assert(backend.includes("var PART_TAG_GROUP_HEADERS = ['Group ID', 'Group Name', 'Prefix', 'Start Number', 'Digits', 'Active', 'Created By', 'Created At', 'Updated By', 'Updated At']"));
assert(backend.includes("var PART_TAG_GROUP_ITEM_HEADERS = ['Group ID', 'Part No', 'Sheet Name', 'Model', 'Part Name', 'Category', 'Added By', 'Added At']"));
// สถานะต้องใช้คำเดียวกับ CWM System ที่หน้างานใช้อยู่ (ติดตั้ง → ถอด)
assert(backend.includes("var PART_TAG_STATUSES = ['รอติดตั้ง', 'ติดตั้งแล้ว', 'ถอดแล้ว', 'ชำรุด/ทิ้ง', 'คืนคลัง']"));
assert(backend.includes("'Installed At', 'Removed At'"), 'ต้องเก็บวันติดตั้ง/วันถอด');
assert(backend.includes("'Log Ref'"), 'ต้องผูกเลขกับแถว Log ที่ออกมัน เพื่อลบทิ้งได้ตอนคืนรายการ');

// ── บังคับเลือกเครื่องตอนเบิกแล้ว จึงถือว่าเบิกไปติดตั้งเลย ไม่ใช่ "รอติดตั้ง" ──
assert(backend.includes('var PART_TAG_STATUS_ON_ISSUE = PART_TAG_STATUS_INSTALLED;'));
assert(/var rowValues = \[[\s\S]{0,700}PART_TAG_STATUS_ON_ISSUE/.test(backend.replace(/\r\n/g, '\n')),
  'ตอนออกเลขต้องใช้สถานะติดตั้งแล้ว ไม่ใช่ PART_TAG_STATUSES[0]');
assert(!/var rowValues = \[[\s\S]{0,700}PART_TAG_STATUSES\[0\]/.test(backend.replace(/\r\n/g, '\n')),
  'ต้องไม่ใช้ค่าเริ่มต้น "รอติดตั้ง" ตอนออกเลขแล้ว');

// ── ของในคลังมีเลขติดอยู่แล้ว → ช่างกรอกเลขที่หยิบมา ระบบไม่ออกเลขเองเป็นหลัก ────
assert(backend.includes('function parsePartTagNosInput(raw)'));
assert(backend.includes('function validatePartTagNosForIssue(payload, pieces)'));
assert(backend.includes('tagNos: e.parameter.tagNos,'), 'ต้องรับ tagNos ผ่าน GET/JSONP ด้วย');
// เลขที่กรอกมาต้องมาก่อนเลขอัตโนมัติเสมอ
assert(backend.includes('var tagNo = suppliedTags[i] || (rule.prefix'),
  'ต้องใช้เลขที่ช่างกรอกก่อน แล้วค่อย fallback ไปออกเลขใหม่');
// ตรวจก่อนแตะ stock — ถ้าเลขผิดต้องล้มทั้งรายการตั้งแต่ยังไม่เขียนอะไร
assert(backend.includes('validatePartTagNosForIssue(payload, Math.abs(signedQty));'));
assert(backend.indexOf('validatePartTagNosForIssue(payload, Math.abs(signedQty));') <
       backend.indexOf('mainSheet.getRange(sheetRowNumber, stockCol + 1).setValue(stockAfter);'),
  'ต้องตรวจเลขก่อนเขียน stock');
// เลขที่ยังติดตั้งอยู่ห้ามถูกเบิกซ้ำ (กรอกผิดชิ้น)
assert(backend.includes('ยังติดตั้งอยู่ที่เครื่อง '));
// ชิ้นเดิมที่ถอด/คืนกลับมาแล้วเบิกใหม่ ต้องอัปเดตแถวเดิม ไม่ใช่สร้างแถวซ้ำเลขเดียวกัน
assert(backend.includes('var rowIndexByTagNo = {};'));
assert(backend.includes('sheet.getRange(existingRow, 1, 1, PART_TAG_HEADERS.length).setValues([rowValues]);'));
// ต้องส่งเลขถัดไปให้หน้าเว็บเสนอได้ สำหรับของที่ยังไม่มีเลขติด
assert(backend.includes('next_tag_no: nextByGroup[it.group_id]'));

// ── กันจับคู่พลาดเพราะ model/name ฟอร์แมตต่างกันข้ามแหล่งข้อมูล (roSearchPool แบบ raw
// vs partsData ที่ผ่าน normalizeRecord) — fallback สุดท้ายเทียบแค่ No.+Sheet ────────
assert(backend.includes("var looseKey = normalizePartTagItemKey(payload.partNo, payload.sheetName, '', '');"));
assert(/function resolvePartTagRule[\s\S]{0,1600}looseKey/.test(backend.replace(/\r\n/g, '\n')),
  'resolvePartTagRule ต้องมี fallback เทียบแค่ No.+Sheet เป็นด่านสุดท้าย');
assert(html.includes('function partTagLooseKeyForItem(item)'));
assert(html.includes('partTagConfigState.byLooseKey'));
assert(/function getPartTagRuleForItem\(item\)[\s\S]{0,400}byLooseKey/.test(html.replace(/\r\n/g, '\n')),
  'getPartTagRuleForItem ต้อง fallback ไปที่ byLooseKey เมื่อคีย์เต็มไม่ตรง');

// ── Regression จริงที่เจอ: สมาชิกที่เพิ่มเข้ากลุ่มตอน Sheet Name ยังบันทึกเป็นค่าว่าง ──
// (เช่นถูกเพิ่มก่อน loadRequestOrderSearchPool จะแก้ให้ติด __sourceSheet มาด้วยเสมอ)
// stored key จึงเป็น "12:::model::name" (sheet ว่าง) แต่ item จริงมี sheet เต็ม "12::coil
// winding::model::name" — ทั้ง byKey (strict) และ byLooseKey (No.+Sheet) ไม่ match เลย
// เพราะฝั่ง loose ก็ยังเทียบ Sheet อยู่ดี ต้องมี tier ที่ไม่สนใจ Sheet เมื่อฝั่งเก็บไว้ว่างจริงๆ
// มิเรอร์ backend's resolvePartTagRule ที่มี tier นี้อยู่แล้ว
assert(html.includes('function partTagEmptySheetKeyForItem(item)'));
assert(html.includes('partTagConfigState.byEmptySheetKey'));
assert(html.includes("if (!String(a.sheet_name || '').trim()) {"),
  'ต้อง index เฉพาะแถวที่ sheet_name ว่างจริงๆ เข้า byEmptySheetKey ไม่ใช่ทุกแถว');
assert(/function getPartTagRuleForItem\(item\)[\s\S]{0,400}byEmptySheetKey/.test(html.replace(/\r\n/g, '\n')),
  'getPartTagRuleForItem ต้อง fallback ไปที่ byEmptySheetKey ก่อน byLooseKey');
// ลำดับ tier ต้องเป็น strict -> empty-sheet -> loose (No.+Sheet) ตามความแม่นยำมากไปน้อย
assert(html.indexOf('partTagConfigState.byKey[partTagKeyForItem(item)]') <
       html.indexOf('byEmptySheetKey || {})[partTagEmptySheetKeyForItem(item)]'));
assert(html.indexOf('byEmptySheetKey || {})[partTagEmptySheetKeyForItem(item)]') <
       html.indexOf('byLooseKey || {})[partTagLooseKeyForItem(item)]'));

// ── คืนรายการต้องลบเลขที่ออกไป เพื่อให้เลขเดิมกลับมาใช้ซ้ำได้ ──────────────────
assert(backend.includes('function deletePartTagsForLogEntry(logTimestamp, partName)'));
assert(backend.includes('var removedTags = deletePartTagsForLogEntry(originalTs, originalName);'),
  'returnLogEntry ต้องลบเลขประจำชิ้นของรายการนั้นด้วย');
assert(backend.includes('removed_tags: removedTags'));
// แถวเก่าที่ยังไม่มี Log Ref ต้องยัง match ได้ ไม่งั้นเลขที่ออกไปก่อนหน้าจะลบไม่ออก
assert(backend.includes('matched = !ref && issuedAt === target'),
  'ต้องมี fallback เทียบเวลาเบิกสำหรับแถวที่ยังไม่มี Log Ref');
// ลบเลขพลาดต้องไม่ทำให้การคืน (ที่คืนสต็อกไปแล้ว) พัง
assert(/function deletePartTagsForLogEntry[\s\S]{0,1400}catch \(err\)/.test(backend.replace(/\r\n/g, '\n')));

// แถว Log กับเลขประจำชิ้นต้องใช้ timestamp ตัวเดียวกัน ไม่งั้นผูกกลับหากันไม่เจอ
assert(backend.includes('var txnTimestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");'));
assert(backend.includes('createPartTagsForIssue(payload, issuedBy, Math.abs(signedQty), txnTimestamp)'));
assert(backend.includes('installedAt = now;') && backend.includes("removedAt = '';"));

// ออกเลขเฉพาะขาเบิกออก (signedQty < 0) — รับเข้าต้องไม่ออกเลข
// (มีเงื่อนไข !payload.skipPartTags ต่อท้ายด้วย เพื่อไม่ออกเลขใหม่ตอน "คืนรายการ")
assert(/if \(signedQty < 0 && !payload\.skipPartTags\) \{/.test(backend));
assert(backend.includes('partTagResult = createPartTagsForIssue(payload, issuedBy, Math.abs(signedQty), txnTimestamp);'));
assert(backend.includes('part_tags: partTagResult.tags'));
assert(backend.includes('part_tag_warning: partTagResult.skipped_reason'));

// เลขเริ่มต้น/จำนวนหลักต่อกลุ่มต้องถูกส่งต่อไปจนถึงจุดออกเลขจริง
assert(backend.includes('nextPartTagRunning(existing, rule.prefix, rule.start_number)'));
// เลขอัตโนมัติเดินตามจำนวนที่ "ออกเองจริง" (autoUsed) ไม่ใช่ index ของลูป
// เพราะบางชิ้นช่างกรอกเลขมาเอง เลขอัตโนมัติต้องไม่ข้ามกระโดด
assert(backend.includes('padPartTagRunning(running + autoUsed, rule.digits)'));
// resolvePartTagRule ต้องรับ payload ทั้งก้อน (จับคู่ระดับรายชิ้น) ไม่ใช่แค่ category
assert(backend.includes('var rule = resolvePartTagRule(payload);'));
assert(!backend.includes('resolvePartTagRule(payload.category)'), 'ต้องไม่ผูกกับ Category ทั้งหมดอีกต่อไป');

// action ต้องถูก dispatch ครบสำหรับ flow กลุ่มแท็ก + ทะเบียนชิ้น
[
  'getPartTagGroups', 'getPartTagAssignments', 'createPartTagGroup', 'updatePartTagGroup',
  'deletePartTagGroup', 'getPartTagGroupItems', 'addPartTagGroupItem', 'removePartTagGroupItem',
  'getPartTags', 'updatePartTagStatus'
].forEach(function (action) {
  assert(backend.includes("action === '" + action + "'"), 'missing dispatch for ' + action);
});
// action category-based เดิมต้องไม่มีเหลืออยู่แล้ว (ถูกแทนที่ทั้งหมด)
assert(!backend.includes("action === 'getPartTagConfig'"));
assert(!backend.includes("action === 'setPartTagConfig'"));

// ── Backend: logic ของฟังก์ชันล้วน (โหลดมารันจริง) ─────────────────────────────
const sandbox = {};
const backendLf = backend.replace(/\r\n/g, '\n');
const pureFns = ['defaultPartTagPrefix', 'nextPartTagRunning', 'padPartTagRunning', 'normalizePartTagItemKey', 'partTagLineFromProcess', 'rowToPartTagGroup', 'rowToPartTagGroupItem'];
pureFns.forEach(function (name) {
  const re = new RegExp('^function[ ]+' + name + '\\([\\s\\S]*?\\n}$', 'm');
  const match = backendLf.match(re);
  assert(match, 'cannot extract ' + name);
  sandbox[name] = new Function('defaultPartTagPrefix', match[0] + '\nreturn ' + name + ';')(sandbox.defaultPartTagPrefix);
});

assert.strictEqual(sandbox.defaultPartTagPrefix('Stripper Cutter'), 'STRI');
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
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'ME'), 1, 'prefix ที่เป็น substring กันต้องไม่ปนกัน (MECH- ไม่ใช่ ME-)');

// "สานต่อ" เลขจากระบบเดิม: prefix ที่ยังไม่เคยออกในทะเบียนนี้เลย ต้องเริ่มจาก startNumber ที่ตั้งไว้
assert.strictEqual(sandbox.nextPartTagRunning([['Tag No']], 'SC', 40), 40, 'ชีทเปล่า + ตั้งเลขเริ่มต้น 40 ต้องเริ่มที่ 40 ไม่ใช่ 1');
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'SC', 40), 40, 'prefix ใหม่ที่ยังไม่มีในทะเบียนต้องใช้เลขเริ่มต้นแม้ prefix อื่นมีข้อมูลอยู่แล้ว');
assert.strictEqual(sandbox.nextPartTagRunning(rows, 'MECH', 999), 8, 'มีเลขอยู่แล้วต้องต่อจาก max ไม่สนใจ startNumber');

// identity key ของอะไหล่ต้องคงที่ (partNo+sheetName+model+partName) trim+lowercase ทั้งหมด
// เพื่อกันสมาชิกกลุ่มไม่ตรงกับตอนเบิกจริงเพราะช่องว่าง/ตัวพิมพ์ต่างกัน
assert.strictEqual(
  sandbox.normalizePartTagItemKey(' SC-01 ', 'Coil Winding', ' ME-1 ', ' Stripper Cutter '),
  sandbox.normalizePartTagItemKey('sc-01', 'coil winding', 'me-1', 'stripper cutter')
);
assert.notStrictEqual(
  sandbox.normalizePartTagItemKey('SC-01', 'Coil Winding', 'ME-1', 'Stripper Cutter'),
  sandbox.normalizePartTagItemKey('SC-01', 'H9', 'ME-1', 'Stripper Cutter'),
  'sheet/line ต่างกันต้องถือเป็นชิ้นคนละตัว (กัน No. ซ้ำข้าม sheet)'
);

// คอลัมน์ Line ต้องไม่มีโน๊ตของการเบิกติดมา
assert.strictEqual(sandbox.partTagLineFromProcess('Coil Winding | note: เปลี่ยนตามรอบ'), 'Coil Winding');
assert.strictEqual(sandbox.partTagLineFromProcess('H9'), 'H9');
assert.strictEqual(sandbox.partTagLineFromProcess(''), '');

const group = sandbox.rowToPartTagGroup(['PTG-1', 'Stripper Cutter', 'SC', '40', '3', 'TRUE', 'admin', '2026-07-30', 'admin', '2026-07-30']);
assert.strictEqual(group.name, 'Stripper Cutter');
assert.strictEqual(group.prefix, 'SC');
assert.strictEqual(group.start_number, 40, 'ต้องอ่านเลขเริ่มต้นที่ admin ตั้งไว้ (สานต่อจาก SC-039 เดิม)');
assert.strictEqual(group.digits, 3, 'ต้องอ่านจำนวนหลักที่ admin ตั้งไว้ (SC-040 = 3 หลัก)');
assert.strictEqual(group.active, true);
assert.strictEqual(sandbox.rowToPartTagGroup(['PTG-2', 'X', 'X', '', '', 'FALSE', 'a', 'b', 'a', 'b']).active, false);
assert.strictEqual(sandbox.rowToPartTagGroup(['PTG-3', 'X', '', '', '', 'TRUE', 'a', 'b', 'a', 'b']).start_number, 1, 'ไม่ระบุเลขเริ่มต้นต้อง fallback เป็น 1');
assert.strictEqual(sandbox.rowToPartTagGroup(['PTG-3', 'X', '', '', '', 'TRUE', 'a', 'b', 'a', 'b']).digits, 5, 'ไม่ระบุจำนวนหลักต้อง fallback เป็น 5');

const groupItem = sandbox.rowToPartTagGroupItem(['PTG-1', 'SC-01', 'Coil Winding', 'ME-1', 'Stripper Cutter', 'Mechanical', 'admin', '2026-07-30']);
assert.strictEqual(groupItem.group_id, 'PTG-1');
assert.strictEqual(groupItem.part_name, 'Stripper Cutter');
assert.strictEqual(groupItem.sheet_name, 'Coil Winding');

// ── Regression: rowToPartTag ต้อง normalize คอลัมน์วัน-เวลาทุกตัว ──────────────
// Sheets แปลงสตริงที่ appendRow เขียนลงเป็น Date object ให้เอง ถ้าอ่านด้วย
// String(dateObj) ตรงๆ จะได้ toString() ดิบ (เช่น มีวงเล็บชื่อเขตเวลาภาษาไทย)
// ซึ่งพัง 2 ทาง: (1) แสดงผลรก (2) ส่งค่านี้กลับไปเป็น timestamp ให้ returnLogEntry
// แล้ว parse ไม่ตรงกับเวลาจริงในแถว Log ทำให้หารายการเบิกต้นทางไม่เจอ คืนของไม่ได้
assert(!/issued_at: String\(r\[10\]/.test(backend), 'issued_at ต้องไม่ใช้ String(...) ดิบแล้ว');
assert(backend.includes('issued_at: normalizeLogTimestamp(r[10]),'));
assert(backend.includes('status_at: normalizeLogTimestamp(r[13]),'));
assert(backend.includes('installed_at: normalizeLogTimestamp(r[16]),'));
assert(backend.includes('removed_at: normalizeLogTimestamp(r[17]),'));
assert(backend.includes('log_ref: normalizeLogTimestamp(r[18])'));

// ── Frontend: UI + การรับเลขกลับมา ────────────────────────────────────────────
assert(html.includes('id="partTagIssuedModal"'));
assert(html.includes('id="partTagRegistryModal"'));
assert(html.includes('id="partTagGroupItemsModal"'));
assert(html.includes('id="partTagGroupList"'));
assert(html.includes('#partTagIssuedModal { z-index: 9800; }'), 'สรุปเลขต้องอยู่เหนือ Issue Cart');
assert(html.includes('if (res.part_tags && res.part_tags.length) {'));
assert(html.includes('showIssuedPartTagsSummary(issuedTagGroups);'));
// ต้อง degrade เงียบๆ ถ้า backend ยังไม่ deploy
assert(html.includes('partTagConfigState.supported = false;'));
assert(html.includes('if (!partTagConfigState.supported || !item) return null;'));
// จับคู่ระดับรายชิ้น (byKey) ไม่ใช่ระดับ Category (byCategory) อีกต่อไป
assert(html.includes('partTagConfigState.byKey[partTagKeyForItem(item)]'));

// ── หน้าเว็บ: ช่องกรอกเลขที่หยิบมา ทั้ง Issue Cart และเบิกด่วน ─────────────────
assert(html.includes('function parseTagNosInput(raw)'));
assert(html.includes('data-cart-tagnos='), 'Issue Cart ต้องมีช่องกรอกเลขต่อรายการ');
assert(html.includes('id="quickIssueTagNos"'), 'เบิกด่วนต้องมีช่องกรอกเลข');
// ค่าที่พิมพ์ต้องไม่หายตอน renderIssueCart วาดใหม่ (เช่นกดเพิ่ม/ลดจำนวน)
assert(html.includes('var issueCartTagNos = {};'));
assert(html.includes('issueCartTagNos[entry.key] = input.value;'));
// จำนวนเลขต้องเท่าจำนวนที่เบิก ทั้งสองฟอร์ม
assert(html.includes("' เลข แต่เบิก ' + wantQty + ' ชิ้น — ต้องใส่ให้ครบทุกชิ้น (คั่นด้วย ,)'"));
assert(html.includes("' เลข แต่เบิก ' + qty + ' ชิ้น — ต้องใส่ให้ครบ (คั่นด้วย ,)'"));
// ส่ง tagNos ไปกับ payload ทั้งสองทาง
assert(html.includes('tagNos: tagNosValue'));
assert(html.includes('tagNos: quickTagNosRaw,'));
assert(html.includes("'tagNos=' + encodeURIComponent(requestPayload.tagNos || ''),"),
  'JSONP query ต้องส่ง tagNos ไปด้วย ไม่งั้น backend ไม่เห็นเลขที่กรอก');
// เบิกสำเร็จต้องล้างเลขที่กรอกไว้ ไม่ให้ค้างไปรอบถัดไป
assert(html.includes('delete issueCartTagNos[key];'));
// แก้จำนวนในเบิกด่วนต้องไม่ล้างเลขที่พิมพ์ไว้
assert(html.includes('function syncQuickIssueTagField(item, resetValue)'));
assert(html.includes('if (resetValue) input.value = \'\';'));
assert(html.includes('syncQuickIssueTagField(item, true);'), 'เปิด modal ใหม่ต้องล้างค่าของชิ้นก่อนหน้า');
assert(!html.includes('byCategory'), 'ต้องไม่เหลือโค้ด config แบบผูก Category เดิม');

// Admin: ช่องตั้งเลขเริ่มต้น (สานต่อจากระบบเดิม เช่น CWM System) ต้องแปลง "040" เป็น start=40, digits=3
assert(html.includes('id="partTagGroupStartInput"'));
assert(html.includes("var startRaw = String((startEl && startEl.value) || '').replace(/\\D/g, '');"));
assert(html.includes("var startNumber = startRaw ? parseInt(startRaw, 10) : 1;"));
assert(html.includes("var digits = startRaw ? startRaw.length : 5;"));

// Admin: เลือกอะไหล่เข้ากลุ่มทีละชิ้นด้วยการค้นหา (ใช้ roSearchPool เดียวกับหน้าขอซื้อ)
assert(html.includes('function searchPartTagGroupItemCandidates()'));
assert(html.includes("action: 'addPartTagGroupItem'"));
assert(html.includes("action: 'removePartTagGroupItem'"));

// ── Regression: backend เก่าตอบ array กลับมา ต้องไม่ถูกอ่านว่า "สำเร็จ" ────────
// Apps Script เวอร์ชันที่ยังไม่ deploy จะไม่รู้จัก action ของทะเบียนชิ้น แล้ว doGet ตกไปที่
// default = คืน array รายการอะไหล่ ซึ่งไม่มี status:'error' จึงหลุด parseApiResponse ไปได้
// ทำให้หน้าเว็บขึ้นว่าสร้างกลุ่มสำเร็จทั้งที่ไม่ได้บันทึกอะไร (อาการที่ผู้ใช้เจอ)
assert(/function parsePartTagResponse\(res, expectedField\)/.test(html));
assert(html.includes('var PART_TAG_DEPLOY_HINT ='));

const guardSrc = html.replace(/\r\n/g, '\n').match(/^ {4}function parsePartTagResponse\([\s\S]*?\n {4}}$/m);
assert(guardSrc, 'cannot extract parsePartTagResponse');
const guardState = { supported: true };
const parsePartTagResponse = new Function(
  'parseApiResponse', 'partTagConfigState', 'PART_TAG_DEPLOY_HINT',
  guardSrc[0] + '\nreturn parsePartTagResponse;'
)(
  function (res) { if (res && res.status === 'error') throw new Error(res.message); return res; },
  guardState,
  'DEPLOY_HINT'
);

// array (backend เก่าตกมาที่ default) ต้อง throw + ปิดฟีเจอร์
assert.throws(function () { parsePartTagResponse([{ no: '1', name: 'part' }], 'group_id'); }, /DEPLOY_HINT/);
assert.strictEqual(guardState.supported, false, 'ต้องปิดฟีเจอร์เมื่อรู้ว่า backend ยังไม่ deploy');
// response ที่ไม่มี field ที่คาดไว้ ก็ต้องไม่ผ่าน
assert.throws(function () { parsePartTagResponse({ status: 'success' }, 'group_id'); }, /DEPLOY_HINT/);
assert.throws(function () { parsePartTagResponse(null, 'groups'); }, /DEPLOY_HINT/);
// response ที่ถูกต้องต้องผ่านและคืนค่าเดิม
guardState.supported = true;
const okRes = { status: 'success', group_id: 'PTG-1' };
assert.strictEqual(parsePartTagResponse(okRes, 'group_id'), okRes);
assert.strictEqual(parsePartTagResponse({ status: 'success', groups: [] }, 'groups').groups.length, 0, 'groups ว่างเป็นค่าที่ถูกต้อง ไม่ใช่ error');
assert.strictEqual(guardState.supported, true);
// error จาก backend จริงต้องยัง throw ตามเดิม (ไม่ถูกกลบด้วย hint)
assert.throws(function () { parsePartTagResponse({ status: 'error', message: 'ไม่มีสิทธิ์' }, 'group_id'); }, /ไม่มีสิทธิ์/);

// Admin ต้องโชว์กล่องเตือน deploy แทน empty state ที่ดูเหมือน "ไม่มีอะไรเกิดขึ้น"
assert(html.includes('if (!partTagConfigState.supported) {'));
assert(html.includes('ยังใช้งานไม่ได้ — ต้อง Deploy Apps Script ก่อน'));

// ── Regression: item จาก roSearchPool ต้องมี __sourceSheet ────────────────────
// Admin เลือกอะไหล่เข้ากลุ่มจาก roSearchPool ถ้า pool ไม่ติดชื่อชีตต้นทางมา จะบันทึก
// sheet_name เป็นค่าว่าง แต่ตอนเบิกจริง payload ส่งชื่อชีตจริงมา → คีย์ไม่ตรงกัน →
// ไม่ออกเลขให้เลยทั้งที่ตั้งค่าถูก (เงียบสนิท หาสาเหตุยาก)
assert(html.includes('item.__sourceSheet = item.__sourceSheet || t.sheet;'),
  'loadRequestOrderSearchPool ต้องติด __sourceSheet ให้ item ในpool');

// backend ต้องมี fallback เทียบแบบไม่สนชีต เฉพาะแถวที่ Sheet Name ว่างจริงๆ
assert(backend.includes("!String(it.sheet_name || '').trim() &&"),
  'fallback ต้องจำกัดเฉพาะแถวที่ Sheet Name ว่าง ไม่ใช่เทียบหลวมทุกแถว');
assert(backend.includes('var keyNoSheet = normalizePartTagItemKey(payload.partNo,'));

// empty state ของทะเบียนชิ้นต้องไม่พูดถึง "Category" อีก (เลิกผูก Category ไปแล้วที่ #540)
// และต้องชี้ทางไปตั้งค่าที่ Admin → กลุ่มแท็กชิ้น ไม่งั้นผู้ใช้เข้าใจว่าระบบเสีย
assert(!html.includes('เลขจะถูกออกอัตโนมัติตอนเบิกอะไหล่ใน Category ที่ตั้งค่าไว้'),
  'empty state ต้องไม่ชี้ให้ไปตั้งค่า Category ที่ไม่มีอยู่แล้ว');
assert(html.includes('Admin → 🏷️ กลุ่มแท็กชิ้น</b>'));

console.log('Part tag registry checks passed');
