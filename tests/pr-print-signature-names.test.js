const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ---- ช่องลายเซ็นใบ PR ที่ปริ้นจาก Inbox: ชื่อจริง + เวลาไทย ----
// เคสจริง: Prepared by ขึ้น "Admin" (username) และ Date ขึ้น "Thu Oct 22 2026 10:14:57 GMT+0700 (เวลาอินโดจีน)"

// 1) Users มีคอลัมน์ display_name — บันทึกจากฟอร์มผู้ใช้ และส่งกลับไปให้หน้าเว็บ
assert(backend.includes("headers.indexOf('display_name')"), 'Users sheet gets a display_name column');
assert(backend.includes("displayName: String(user.displayName || '')"), 'sanitizeUserForClient exposes displayName');
assert(html.includes("displayName: String((document.getElementById('uDisplayName') || {}).value || '').trim()"), 'upsertUser sends displayName');

// 2) getPRForApproval ส่งชื่อจริงของผู้ขอ/ผู้อนุมัติ และเวลาแบบ normalize (Asia/Bangkok)
const start = backend.indexOf('function getPRForApproval(payload)');
const body = backend.slice(start, backend.indexOf('\nfunction ', start + 10));
assert(body.includes('header.created_by_name ='), 'returns created_by_name');
assert(body.includes('header.approved_by_name ='), 'returns approved_by_name');
assert(body.includes('header.approved_at = normalizeLogTimestamp('), 'approved_at normalized');

// 3) หน้าปริ้นใช้ชื่อจริงก่อน username และแปลงเวลา Date ดิบเป็นเวลาไทย
assert(html.includes("preparedByOverride: header.created_by_name || header.created_by || '-'"), 'prepared by uses display name');
assert(html.includes("approvedByOverride: header.approved_by_name || header.approved_by || ''"), 'approved by uses display name');
const m = html.match(/function formatThaiDateTime\(value\) \{[\s\S]*?\n    \}/);
assert(m, 'formatThaiDateTime exists');
const formatThaiDateTime = new Function(m[0] + '; return formatThaiDateTime;')();
const label = formatThaiDateTime('Thu Oct 22 2026 10:14:57 GMT+0700 (เวลาอินโดจีน)');
assert.strictEqual(label.slice(0, 16), '22/10/2026 10:14', 'raw Date string → Thai time');
assert.strictEqual(formatThaiDateTime('Thu Oct 22 2026 03:14:57 GMT+0000').slice(0, 16), '22/10/2026 10:14', 'UTC → Bangkok');

console.log('pr-print-signature-names: ok');

// 4) ตัดแถว Checked by ออกจากตารางลายเซ็น (เหลือ Prepared by / Approved by)
assert(!html.includes('<tr><td>Checked by</td>'), 'Checked by row removed');
