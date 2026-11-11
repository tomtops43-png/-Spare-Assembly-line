const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// หน้าติดตาม PR: ป็อปอัปรายละเอียดต้องมีปุ่มเปิดใบ PR ฉบับเต็ม (เคสจริง: เปิดดูรายละเอียด PR ไม่ได้)
assert(html.includes("var PT_DOC_VIEW_STATUSES = ['APPROVED', 'ORDERED', 'PARTIAL', 'RECEIVED', 'CLOSED'];"), 'เปิดได้ทุกสถานะหลังอนุมัติ');
const render = html.slice(html.indexOf('function ptRenderDetail()'), html.indexOf('function ptDetailAction('));
assert(render.includes('data-pt-act="view-doc"') && render.includes('📄 ดูใบ PR'), 'มีปุ่ม ดูใบ PR');
assert(render.includes('PT_DOC_VIEW_STATUSES.indexOf(h.status) > -1'), 'ปุ่มขึ้นตามสถานะ');
const action = html.slice(html.indexOf('function ptDetailAction('), html.indexOf('function initPrTrackingPage'));
assert(/act === 'view-doc'[\s\S]*getPRForApproval[\s\S]*openApprovedPrPrintWindow\(res\.header, res\.lines\)/.test(action), 'กดแล้วโหลด PR และเปิดใบเต็ม');

console.log('pr-tracking-view-doc: ok');
