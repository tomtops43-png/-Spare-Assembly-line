const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

function grab(re, label) {
  const m = backend.replace(/\r\n/g, '\n').match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── เทียบค่าเก่า/ใหม่ ต้องไม่สร้าง log ปลอม ────────────────────────────────────
// ค่าจากชีตเป็น number ('10') แต่จากฟอร์มเป็น string ("10") — ถ้าเทียบ String ตรงๆ
// จะขึ้นว่า "แก้ Min 10 → 10" ทุกครั้งที่กดบันทึก จน log จริงจมหาย
const eqSrc = grab(/^function itemAuditValueEquals\(a, b\) \{[\s\S]*?\n\}/m, 'itemAuditValueEquals');
const itemAuditValueEquals = new Function(eqSrc + '\nreturn itemAuditValueEquals;')();
assert.strictEqual(itemAuditValueEquals(10, '10'), true, 'number กับ string ตัวเลขเดียวกันต้องถือว่าไม่เปลี่ยน');
assert.strictEqual(itemAuditValueEquals('10', '10.0'), true, '10 กับ 10.0 คือค่าเดียวกัน');
assert.strictEqual(itemAuditValueEquals(' 5 ', 5), true, 'ช่องว่างหัวท้ายไม่ใช่การเปลี่ยนค่า');
assert.strictEqual(itemAuditValueEquals('', null), true);
assert.strictEqual(itemAuditValueEquals(undefined, ''), true);
assert.strictEqual(itemAuditValueEquals(10, '2'), false, 'ลด Min จาก 10 เหลือ 2 ต้องถูกจับได้');
assert.strictEqual(itemAuditValueEquals('SKF', 'NSK'), false);
assert.strictEqual(itemAuditValueEquals('', '5'), false, 'จากว่างเป็นมีค่า = เปลี่ยน');
// ข้อความที่ไม่ใช่ตัวเลขห้ามถูกแปลงเป็น NaN แล้วเทียบว่าเท่ากัน
assert.strictEqual(itemAuditValueEquals('abc', 'xyz'), false);

// ── ปิดประตูหลังของยอดคงเหลือ ─────────────────────────────────────────────────
// ฟอร์มแก้ไขอะไหล่ส่ง stock มาด้วยเสมอ ถ้าเขียนทับตรงๆ = ปรับยอดได้โดยไม่ลง Log
// ไม่ต้องนับสต็อก ไม่ต้องมีเหตุผล ซึ่งขัดกับกติกา "ปรับยอดได้เฉพาะ Admin ผ่านการนับ"
assert(/var actorRole = normalizeRole\(payload\.actorRole\);/.test(backend));
assert(/if \(nextStockText === '' \|\| itemAuditValueEquals\(oldStockValue, nextStockText\)\) \{[\s\S]{0,80}stockLocked = true;/.test(backend),
  'ค่าเท่าเดิมหรือว่าง = ไม่เขียนทับ (กัน false-block ตอนข้อมูลหน้าเว็บเก่ากว่าชีต)');
assert(/\} else if \(actorRole !== 'admin'\) \{[\s\S]{0,120}stockLocked = true;/.test(backend), 'ไม่ใช่ Admin = ไม่เขียนยอดคงเหลือ');
assert(/if \(key === 'stock' && stockLocked\) continue;/.test(backend));
assert(backend.indexOf('ปรับยอดต้องผ่าน รับเข้า/เบิกออก หรือการตรวจนับสต็อก (Admin เท่านั้น)') > -1,
  'ต้องบอกกลับไปด้วยว่าทำไมยอดไม่เปลี่ยน ไม่ใช่เงียบ');
// Admin แก้ได้แต่ต้องเด้งเป็นประเภทแยกให้เห็นชัด ไม่ปนกับการแก้ข้อมูลทั่วไป
assert(/action: auditChanges\.stock \? 'STOCK_EDIT' : 'UPDATE',/.test(backend));

// ── ต้องรู้ว่าใครแก้ ────────────────────────────────────────────────────────
assert((backend.match(/var itemActor = requirePermission\(authPayload, 'manage_items'\);/g) || []).length === 2,
  'ทั้ง GET และ POST ต้องส่งชื่อผู้แก้ไขเข้าไป');
assert((backend.match(/actor: itemActor\.username,/g) || []).length === 2);
assert(/actor: delActor\.username/.test(backend) && /actor: delActorPost\.username/.test(backend), 'การลบก็ต้องรู้ว่าใครลบ');

// ── ครอบคลุมทั้งเพิ่ม / แก้ / ลบ ─────────────────────────────────────────────
assert(/action: 'CREATE', changes: createChanges/.test(backend));
assert(/action: 'DELETE', changes: deleteChanges/.test(backend));
// ลบแล้วชีตไม่เหลือแถวให้ดู ต้องเก็บภาพก่อนลบ ไม่ใช่หลังลบ
assert(/var snapshot = itemAuditSnapshotFromRow\(ctx, ctx\.rows\[i\]\);[\s\S]{0,600}ctx\.sheet\.deleteRow\(rowNumber\);/.test(backend),
  'ต้องเก็บ snapshot ก่อนเรียก deleteRow');
// เขียน log พังต้องไม่ทำให้การแก้ข้อมูลที่สำเร็จแล้วกลายเป็น error
assert(/function appendItemAudit\(entry\) \{[\s\S]{0,1400}catch \(err\) \{[\s\S]{0,160}Logger\.log\('appendItemAudit warning/.test(backend));
assert(backend.includes("var ITEM_AUDIT_HEADERS = ['Date Time', 'User', 'Role', 'Sheet Name', 'NO', 'Part Name', 'Model', 'Action Type', 'Changed Fields', 'Old Values', 'New Values'];"));
assert(backend.includes("SPARE_APP_CONFIG.itemAuditSheetName = SPARE_APP_CONFIG.itemAuditSheetName || 'ItemAudit'"));
assert(backend.includes("if (action === 'getItemAudit') return respond(getItemAudit(e.parameter), e);"));
assert(backend.includes("if (action === 'getItemAudit') return respond(getItemAudit(body), e);"));
// ชีตนี้โตเรื่อยๆ ห้ามส่งทั้งก้อนกลับไป
assert(/for \(var i = values\.length - 1; i >= 1 && rows\.length < limit; i -= 1\)/.test(backend), 'ต้องอ่านใหม่สุดก่อนและตัดตาม limit');

// ── หน้าจอฝั่ง Admin ────────────────────────────────────────────────────────
assert(htmlLf.includes('id="itemAuditTable"') && htmlLf.includes('🛡️ ประวัติการแก้ไขข้อมูลอะไหล่'));
assert(htmlLf.includes('id="itemAuditActionFilter"') && htmlLf.includes('id="itemAuditSearch"'));
assert(/setTabStyles\('admin'\);\n\s+loadItemAudit\(false\)/.test(htmlLf), 'เข้าหน้า Admin แล้วต้องโหลดประวัติให้เลย');
assert(/STOCK_EDIT: \{ label: '⚠️ แก้ยอดคงเหลือ'/.test(htmlLf), 'การแก้ยอดคงเหลือต้องเด่นกว่าการแก้อย่างอื่น');
assert(htmlLf.includes('⚠️ มีการแก้ยอดคงเหลือโดยตรง '), 'ต้องสรุปให้เห็นว่ามีการแก้ยอดตรงๆ กี่ครั้ง');
// ช่อง Stock ในฟอร์มแก้ไขต้องล็อกให้คนที่ไม่ใช่ Admin เห็นตั้งแต่แรก ไม่ใช่ให้พิมพ์แล้วค่อยเด้ง
assert(/var eStockLocked = !isAdminUser\(\);/.test(htmlLf) && /eStockEl\.readOnly = eStockLocked;/.test(htmlLf));
assert(htmlLf.includes('🔒 ปรับยอดได้เฉพาะ Admin'));
// backend ปฏิเสธแล้วต้องเด้งบอก ไม่ใช่บันทึกผ่านแบบเงียบๆ
assert(/if \(res && res\.warning && typeof showQuickToast === 'function'\)/.test(htmlLf));

console.log('Item audit trail checks passed');
