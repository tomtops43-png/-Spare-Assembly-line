// preset "ผู้บริหาร / Executive Viewer" — ดู Dashboard / Log / Export ได้ทุกไลน์
// แต่ต้องแตะระบบไม่ได้เลย
//
// เทสต์นี้ไม่ได้เช็คแค่ว่ามี preset อยู่ แต่ "รันโค้ดจริง" สองชั้นต่อกันเพื่อพิสูจน์ผลลัพธ์:
//   ชั้น 1  syncPermissionsJsonFromUI() ของฝั่งเว็บ → ได้ permissions_json { allow, deny }
//   ชั้น 2  getRoleDefaultPermissions() + mergePermissions() ของฝั่งเซิร์ฟเวอร์ → สิทธิ์สุดท้าย
// แล้วยืนยันว่าสิทธิ์สุดท้ายไม่มีตัวที่เขียนข้อมูลได้เลยแม้แต่ตัวเดียว
//
// จุดที่พลาดง่ายและเคยเป็นช่องจริง (กันไว้ในเทสต์นี้):
//   ก) role default ของ 'user' เปิด transact / pr_create / request_order_create ให้เอง
//      preset ที่ list แต่ allow จึงปิดไม่ได้ ต้องมี deny ด้วย
//   ข) hasPermission() มี role fallback ของ request_order_* ที่ชนะ deny
//   ค) 5 ฟังก์ชันฝั่งเซิร์ฟเวอร์ที่ "เขียนข้อมูล" ใช้ view_logs เป็นด่านเดียว
//      ซึ่งผู้บริหารต้องมีเพื่อเปิดหน้า Log/Dashboard
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];

function grabFrontFn(name) {
  const re = new RegExp('^ {4}function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n {4}\\}', 'm');
  const m = script.match(re);
  assert(m, 'ต้องดึงฟังก์ชัน ' + name + ' ออกมาได้');
  return m[0];
}
function grabBackFn(name) {
  const re = new RegExp('^function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n\\}', 'm');
  const m = backend.replace(/\r\n/g, '\n').match(re);
  assert(m, 'ต้องดึงฟังก์ชัน backend ' + name + ' ออกมาได้');
  return m[0];
}

// ── โครงสร้าง preset ────────────────────────────────────────────────────
const presetsSrc = script.match(/var ROLE_PRESETS = \{[\s\S]*?\n {4}\};/)[0];
const ROLE_PRESETS = vm.runInNewContext('(' + presetsSrc.replace(/^var ROLE_PRESETS = /, '').replace(/;$/, '') + ')');
const P = ROLE_PRESETS.viewer_exec;
assert(P, 'ต้องมี preset viewer_exec (ผู้บริหาร)');
assert.strictEqual(P.role, 'user', 'ต้องเป็น role user — role leader/admin มี fallback ที่ให้สิทธิ์อนุมัติมาเอง');
assert.strictEqual(P.lines, 'view', 'ต้องเห็นทุกไลน์แบบ view (ห้าม managed ไม่งั้นแก้ข้อมูลไลน์ได้)');
['view_stock', 'view_detail', 'view_alerts', 'view_logs', 'view_dashboard', 'export_data'].forEach(function(k) {
  assert(P.permissions.indexOf(k) > -1, 'preset ผู้บริหารต้องมีสิทธิ์ดู: ' + k);
});
// ห้ามมีสิทธิ์ที่ทำอะไรกับระบบได้
['issue_part', 'receive_part', 'use_issue_cart', 'create_item', 'edit_item', 'delete_item',
 'manage_minmax', 'manage_item_image', 'delete_logs', 'manage_users', 'delete_users',
 'assign_permissions', 'access_admin_panel', 'reset_password',
 'request_order_create', 'request_order_approve', 'request_order_reject',
 'request_order_convert_pr', 'request_order_close', 'request_order_edit',
 'manage_h9', 'manage_arc_chute', 'manage_coil_winding', 'manage_lug_screw'].forEach(function(k) {
  assert(P.permissions.indexOf(k) === -1, 'preset ผู้บริหารต้องไม่มีสิทธิ์: ' + k);
});
// สิทธิ์ที่ไม่มี checkbox ต้องถูกสั่งไว้ให้ครบ
assert(Array.isArray(P.extraDeny), 'ต้องมี extraDeny เพราะ role default เปิด pr_create ให้เอง');
['pr_create', 'pr_approve', 'transact'].forEach(function(k) {
  assert(P.extraDeny.indexOf(k) > -1, 'extraDeny ต้องมี ' + k);
});
assert(Array.isArray(P.extraAllow) && P.extraAllow.indexOf('pr_view_all') > -1,
  'ผู้บริหารควรดูสถานะ PR ได้ (pr_view_all) แต่ต้องไม่มี pr_approve');

// ── ชั้น 1: รัน syncPermissionsJsonFromUI() ตัวจริงด้วย DOM จำลอง ────────
// จำลองเฉพาะที่ฟังก์ชันนี้เรียกใช้ — checkbox สิทธิ์, checkbox ไลน์, ช่อง hidden
const PERMISSION_KEYS = (function() {
  const groupsSrc = script.slice(script.indexOf('var PERMISSION_GROUPS = ['));
  const end = groupsSrc.indexOf("key: 'system'");
  const seg = groupsSrc.slice(0, groupsSrc.indexOf('];', end));
  const out = [];
  const re = /\{ key: '([a-z_0-9]+)', label:/g;
  let m;
  while ((m = re.exec(seg))) out.push(m[1]);
  return out;
})();
assert(PERMISSION_KEYS.length > 20, 'ต้องอ่านรายการ checkbox สิทธิ์ออกมาได้');

function buildUi(preset) {
  const permBoxes = PERMISSION_KEYS.map(function(key) {
    return { __key: key, checked: preset.permissions.indexOf(key) > -1, getAttribute: function() { return key; } };
  });
  const lineBoxes = {};
  const LINE_KEYS = ['h9', 'arc_chute', 'coil_winding', 'lug_screw'];
  LINE_KEYS.forEach(function(k) {
    lineBoxes['[data-line-view="' + k + '"]'] = { checked: preset.lines === 'view' || preset.lines === 'managed' };
    lineBoxes['[data-line-manage="' + k + '"]'] = { checked: preset.lines === 'managed' };
  });
  const hidden = { value: '' };
  const doc = {
    querySelectorAll: function(sel) {
      if (sel === '[data-permission-key]') return permBoxes;
      if (sel === '[data-permission-key]:checked') return permBoxes.filter(function(b) { return b.checked; });
      if (sel === '[data-line-view],[data-line-manage]') return Object.keys(lineBoxes).map(function(k) { return lineBoxes[k]; });
      return [];
    },
    querySelector: function(sel) { return lineBoxes[sel] || null; },
    getElementById: function(id) { return id === 'uPermissionsJson' ? hidden : null; }
  };
  return { doc: doc, hidden: hidden };
}

function runSync(preset) {
  const ui = buildUi(preset);
  const src = [
    script.match(/var PERMISSION_COMPAT_MAP = \{[\s\S]*?\n {4}\};/)[0],
    script.match(/var LINE_OPTIONS = \[[\s\S]*?\n {4}\];/)[0],
    'var uAllLinesAccess = { checked: ' + (preset.lines === 'all') + ' };',
    'var permissionSummary = null;',
    'var presetExtraAllow = ' + JSON.stringify(preset.extraAllow || []) + ';',
    'var presetExtraDeny = ' + JSON.stringify(preset.extraDeny || []) + ';',
    grabFrontFn('getSelectedPermissionsFromUI'),
    grabFrontFn('getSelectedLineAccessFromUI'),
    grabFrontFn('getManagedPermissionKeyMap'),
    grabFrontFn('syncPermissionsJsonFromUI'),
    'syncPermissionsJsonFromUI();',
    'return JSON.parse(document.getElementById("uPermissionsJson").value);'
  ].join('\n');
  return new Function('document', src)(ui.doc);
}

const execJson = runSync(P);
assert(Array.isArray(execJson.allow) && Array.isArray(execJson.deny), 'ต้องได้ { allow, deny }');
// สิทธิ์ที่ไม่มี checkbox ต้องไหลลงไปใน json จริง
assert(execJson.allow.indexOf('pr_view_all') > -1, 'extraAllow ต้องอยู่ใน allow');
assert(execJson.allow.indexOf('view_purchase_history') > -1, 'extraAllow ต้องอยู่ใน allow');
assert(execJson.deny.indexOf('pr_create') > -1, 'extraDeny ต้องอยู่ใน deny');
// ติ๊ก issue_part/receive_part ออก ต้องทำให้ transact ถูก deny ด้วย (ไม่งั้น backend ยังให้เบิกได้)
assert(execJson.deny.indexOf('transact') > -1, 'ปิดเบิก/รับเข้า ต้อง deny transact ด้วย');
assert(execJson.deny.indexOf('manage_items') > -1, 'ปิดแก้ไขรายการ ต้อง deny manage_items ด้วย');
// เห็นทุกไลน์แบบ view เท่านั้น
['view_h9', 'view_arc_chute', 'view_coil_winding', 'view_lug_screw'].forEach(function(k) {
  assert(execJson.allow.indexOf(k) > -1, 'ต้องเห็นไลน์: ' + k);
});
['manage_h9', 'manage_arc_chute', 'manage_coil_winding', 'manage_lug_screw'].forEach(function(k) {
  assert(execJson.allow.indexOf(k) === -1, 'ต้องไม่มีสิทธิ์จัดการไลน์: ' + k);
});

// ── ชั้น 2: merge กับ role default ฝั่งเซิร์ฟเวอร์ → สิทธิ์สุดท้าย ────────
const merged = (function() {
  const src = [
    grabBackFn('normalizeRole'),
    grabBackFn('getRoleDefaultPermissions'),
    grabBackFn('mergePermissions'),
    'return mergePermissions(getRoleDefaultPermissions("user"), custom);'
  ].join('\n');
  return new Function('custom', src)({ allow: execJson.allow, deny: execJson.deny });
})();

// ต้องดูได้
['view', 'view_logs', 'view_dashboard', 'export_data'].forEach(function(k) {
  assert.strictEqual(merged[k], true, 'ผู้บริหารต้องมีสิทธิ์ดู: ' + k);
});
// ต้องเขียนอะไรไม่ได้เลย — นี่คือหัวใจของเทสต์นี้
[
  'transact', 'issue_part', 'receive_part', 'use_issue_cart',
  'manage_items', 'delete_items', 'create_item', 'edit_item', 'delete_item',
  'manage_minmax', 'manage_item_image',
  'pr_create', 'pr_approve',
  'request_order_create', 'request_order_approve', 'request_order_reject',
  'request_order_convert_pr', 'request_order_close', 'request_order_edit',
  'delete_logs', 'manage_users', 'delete_users', 'manage_auth', 'add_user', 'delete_user'
].forEach(function(k) {
  assert.notStrictEqual(merged[k], true,
    'สิทธิ์สุดท้ายต้องไม่เปิด "' + k + '" — ไม่งั้นบัญชีผู้บริหารแตะระบบได้');
});
// role default ของ 'user' เปิด transact/pr_create ให้เอง — ยืนยันว่า merge แล้วถูกปิดจริง
const rawDefaults = (function() {
  const src = [grabBackFn('normalizeRole'), grabBackFn('getRoleDefaultPermissions'), 'return getRoleDefaultPermissions("user");'].join('\n');
  return new Function(src)();
})();
assert.strictEqual(rawDefaults.transact, true, 'ยืนยันสมมติฐาน: role user เปิด transact มาโดยค่าเริ่มต้น');
assert.strictEqual(rawDefaults.pr_create, true, 'ยืนยันสมมติฐาน: role user เปิด pr_create มาโดยค่าเริ่มต้น');

// เทียบกับ Viewer ธรรมดา — ต้องต่างกันที่ "ดู Log/Dashboard/Export" เท่านั้น
const plainJson = runSync(ROLE_PRESETS.viewer);
assert(plainJson.allow.indexOf('view_logs') === -1, 'Viewer ธรรมดาต้องไม่เห็น Log');
assert(execJson.allow.indexOf('view_logs') > -1, 'ผู้บริหารต้องเห็น Log');
assert(execJson.allow.indexOf('view_dashboard') > -1, 'ผู้บริหารต้องเห็น Dashboard');

// ── ปิดช่อง ก) deny ต้องชนะ role fallback ของ request_order_* ────────────
const hasPermSrc = grabFrontFn('hasPermission');
assert(/isPermissionDeniedExplicitly\(permissionName\)/.test(hasPermSrc),
  'hasPermission ต้องเช็ค deny ที่ตั้งใจปิดก่อนใช้ role fallback ของ request_order_*');
assert(hasPermSrc.indexOf('isPermissionDeniedExplicitly') < hasPermSrc.indexOf('hasRequestOrderRoleFallback'),
  'ต้องเช็ค deny ก่อน ไม่ใช่หลัง');
const denyFnSrc = grabFrontFn('isPermissionDeniedExplicitly');
assert(/permissionsJson/.test(denyFnSrc),
  'ต้องอ่านจาก permissions_json ตัวจริง — currentUser.permissions ที่ merge แล้วแยกไม่ออกว่า false นั้นตั้งใจปิดหรือ default ไม่เปิด');
assert(/denySetCacheKey/.test(denyFnSrc), 'ควร cache ไว้ เพราะ hasPermission ถูกเรียกถี่มาก');

// รันจริง: role 'user' ที่ถูก deny request_order_create ต้องได้ false
const hasPermRun = (function() {
  const src = [
    script.match(/var PERMISSION_ALIASES = \{[\s\S]*?\n {4}\};/)[0],
    'var currentUser = null;',
    'var denySetCacheKey = null; var denySetCache = {};',
    grabFrontFn('parsePermissionsJsonSafe'),
    grabFrontFn('isPermissionDeniedExplicitly'),
    grabFrontFn('hasRequestOrderRoleFallback'),
    grabFrontFn('hasPermission'),
    'return function(user, key) { currentUser = user; return hasPermission(key); };'
  ].join('\n');
  return new Function(src)();
})();
const execUser = {
  role: 'user',
  permissions: merged,
  permissionsJson: JSON.stringify({ allow: execJson.allow, deny: execJson.deny })
};
assert.strictEqual(hasPermRun(execUser, 'request_order_create'), false,
  'ปุ่ม "ขอซื้อ" ต้องไม่โผล่กับบัญชีผู้บริหาร — role fallback ต้องไม่ชนะ deny');
assert.strictEqual(hasPermRun(execUser, 'request_order_approve'), false, 'ต้องอนุมัติคำขอไม่ได้');
assert.strictEqual(hasPermRun(execUser, 'request_order_view_all'), true, 'แต่ต้องดูคำขอทั้งหมดได้');
assert.strictEqual(hasPermRun(execUser, 'view_logs'), true, 'ต้องดู Log ได้');
assert.strictEqual(hasPermRun(execUser, 'issue_part'), false, 'ต้องเบิกไม่ได้');
// ช่างทั่วไปต้องไม่ได้รับผลกระทบจากการแก้ fallback
const techJson = runSync(ROLE_PRESETS.user);
const techMerged = (function() {
  const src = [grabBackFn('normalizeRole'), grabBackFn('getRoleDefaultPermissions'), grabBackFn('mergePermissions'),
    'return mergePermissions(getRoleDefaultPermissions("user"), custom);'].join('\n');
  return new Function('custom', src)({ allow: techJson.allow, deny: techJson.deny });
})();
assert.strictEqual(hasPermRun({ role: 'user', permissions: techMerged, permissionsJson: JSON.stringify(techJson) }, 'request_order_create'), true,
  'ช่างทั่วไปต้องยังขอซื้อได้ — การแก้ fallback ต้องไม่กระทบคนเดิม');

// ── ปิดช่อง ค) ฟังก์ชันที่เขียนข้อมูลต้องไม่ใช้ view_logs เป็นด่านเดียว ──
const writers = ['createPurchaseHistoryBatch', 'addManualPurchaseHistory', 'uploadPurchaseHistoryAttachment',
  'saveStockCountResult', 'importLegacyStockCount'];
writers.forEach(function(fn) {
  const at = backend.indexOf('function ' + fn + '(');
  assert(at > -1, 'ต้องมีฟังก์ชัน ' + fn);
  const body = backend.slice(at, at + 500);
  assert(/requireWarehouseWriter/.test(body),
    fn + ' เขียนข้อมูลลงชีต ต้องผ่านด่าน requireWarehouseWriter — ใช้ view_logs เป็นด่านเดียวไม่พอ ' +
    'เพราะบัญชีดูอย่างเดียวต้องมี view_logs เพื่อเปิดหน้า Log/Dashboard');
});
const gateSrc = grabBackFn('requireWarehouseWriter');
assert(/normalizeRole\(user\.role\) === 'admin'/.test(gateSrc), 'Admin ต้องผ่านด่านนี้เสมอ');
['receive_part', 'issue_part', 'transact', 'manage_items'].forEach(function(k) {
  assert(gateSrc.indexOf("'" + k + "'") > -1, 'ด่านต้องยอมรับสิทธิ์ทำงานคลัง: ' + k);
});
assert(/throw new Error/.test(gateSrc), 'บัญชีดูอย่างเดียวต้องถูกปฏิเสธ');
assert(/ดูอย่างเดียว/.test(gateSrc), 'ข้อความ error ต้องบอกสาเหตุที่ผู้ใช้เข้าใจ');

// ── ปิดช่อง: หน้า/ปุ่มที่บันทึกข้อมูลต้องไม่โผล่กับบัญชีดูอย่างเดียว ──────
const styleSrc = grabFrontFn('setTabStyles');
assert(/tabStockCount\.classList\.toggle\('hidden', !\(hasPermission\('view_logs'\) && canWriteWarehouse\(\)\)\)/.test(styleSrc),
  'หน้าเช็คสต็อกเป็นงานบันทึกข้อมูล ต้องซ่อนจากบัญชีดูอย่างเดียว (ไม่ใช่เปิดตาม view_logs เพียวๆ)');
const canWriteSrc = grabFrontFn('canWriteWarehouse');
['receive_part', 'issue_part', 'manage_items'].forEach(function(k) {
  assert(canWriteSrc.indexOf("'" + k + "'") > -1, 'canWriteWarehouse ต้องเช็ค ' + k + ' ให้ตรงกับด่านฝั่งเซิร์ฟเวอร์');
});
assert(/isAdminUser\(\)/.test(canWriteSrc), 'Admin ต้องผ่าน');
// สร้าง PR ต้องเช็ค pr_create เอง ไม่พึ่ง role
assert(/if \(!hasPermission\('pr_create'\)\) \{/.test(script),
  'ปุ่มสร้าง PR ต้องเช็ค pr_create ก่อน — role default ของ user เปิดสิทธิ์นี้ให้เอง');

// ── การ์ด preset ต้องมีอยู่บนหน้าจอจริง ─────────────────────────────────
assert(html.indexOf('data-role-preset="viewer_exec"') > -1, 'ต้องมีการ์ด preset ผู้บริหารในหน้า Admin');
assert(/ผู้บริหาร/.test(html.slice(html.indexOf('data-role-preset="viewer_exec"'), html.indexOf('data-role-preset="viewer_exec"') + 600)),
  'การ์ดต้องมีคำว่า "ผู้บริหาร" ให้ Admin เลือกถูกใบ');

console.log('role-preset-executive-viewer: OK (allow ' + execJson.allow.length + ' · deny ' + execJson.deny.length + ')');
