// ปุ่ม "📋 แปลงเป็น PR" ในหน้าขอซื้ออะไหล่ — เดิมแค่เปลี่ยนป้ายสถานะ ไม่สร้าง PR จริง ไม่มีอะไรเข้า Inbox
// (และ backend รับ action นี้แค่ทาง POST ส่วนหน้าเว็บส่งทาง GET/JSONP → กดแล้วขึ้น ✅ แต่ไม่เกิดอะไร)
// ตอนนี้: กดแล้วไปหน้า PR Report โหมด "PR ขอซื้ออะไหล่" พร้อมรายการนั้น → ส่งอนุมัติ → เข้า Inbox + ผูกคำขอกับเลข PR
const { html, grabFn, assert } = require('./_export-extract');
const fs = require('fs');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

// ---- หน้าเว็บ: openPrFromRequests พาไปหน้า PR พร้อมรายการ แล้ว Generate ให้ ----
const open = grabFn('openPrFromRequests');
const calls = [];
const win = {};
const env = {
  window: win,
  showQuickToast: function(m, t) { calls.push(['toast', t]); },
  hasPermission: function(p) { return p === 'pr_create'; },
  closeRequestOrderDetailModal: function() { calls.push(['close']); },
  switchToPrReportPage: function() { calls.push(['switch']); },
  prMode: { value: 'assembly_monthly' },
  document: { getElementById: function(id) { return id === 'prRunBtn' ? { click: function() { calls.push(['run']); } } : null; } },
  runPrCandidateSelection: function() { calls.push(['runDirect']); }
};
win.__applyPrPreset = function(m) { calls.push(['preset', m]); };
const fn = new Function(Object.keys(env).join(','), open + '\nreturn openPrFromRequests;').apply(null, Object.keys(env).map(function(k) { return env[k]; }));
fn([{ request_id: 'REQ-9', item_name: 'เครื่องเชื่อมจุดไฟ' }]);
assert.deepStrictEqual(win.__prFromRequests.map(function(r) { return r.request_id; }), ['REQ-9'], 'ส่งรายการที่เลือกไปให้หน้า PR');
assert.deepStrictEqual(calls.filter(function(c) { return c[0] !== 'toast'; }), [['close'], ['switch'], ['preset', 'from_request'], ['run']],
  'ปิด modal → ไปหน้า PR → โหมด PR ขอซื้ออะไหล่ → Generate');
calls.length = 0; win.__prFromRequests = null;
fn([]);
assert.strictEqual(win.__prFromRequests, null, 'ไม่มีรายการ = ไม่ไปไหน');

// ปุ่มในการ์ดคำขอเรียก openPrFromRequests แทนการเปลี่ยนป้ายสถานะเฉยๆ
assert(/if \(act === 'convert'\) \{[\s\S]{0,400}openPrFromRequests\(convRows\.slice\(0, 1\)\);\s*return;/.test(html), 'ปุ่มแปลงเป็น PR เปิดหน้าทำ PR จริง');
assert(!html.includes("requestApi(withAuthPayload({ action: 'convertOrderRequestsToPR', request_ids: [id] }))"), 'เลิกเปลี่ยนป้ายเฉยๆ แล้ว');
// ปุ่ม "สร้าง PR Report ประจำเดือน" ก็ต้องตั้งโหมดให้ด้วย (เดิมแค่สลับแท็บ รายการไม่ถูกดึงมา)
assert(html.includes('openPrFromRequests(selItems);'));

// ส่ง PR แล้วแนบเลขคำขอต้นทางไปให้ backend ผูก
const submit = html.slice(html.indexOf('function prSubmitCreatePr()'), html.indexOf('function prSubmitCreatePr()') + 8000);
assert(submit.includes("payload.source_request_ids = sourceRequestIds.join(',');"));

// ---- Backend ----
assert(backend.includes("if (action === 'convertOrderRequestsToPR') return respond(convertOrderRequestsToPR(e.parameter), e);"), 'GET รับ action นี้ด้วย');
const parse = new Function(grabFn('parseOrderRequestIdList', backend, 0) + '\nreturn parseOrderRequestIdList;')();
assert.deepStrictEqual(parse(['A', ' B ']), ['A', 'B']);
assert.deepStrictEqual(parse('A,B,'), ['A', 'B'], 'GET ส่งมาเป็นสตริงคั่น ,');
assert.deepStrictEqual(parse(undefined), []);

console.log('order-request-convert-to-pr: OK');
