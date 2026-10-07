// ระบบ PR → สั่งซื้อ → รับของ (GR) — รันฟังก์ชันจริงของ Backend.gs บนชีตจำลองในหน่วยความจำ
// ครอบคลุม: ก่อนเปิดระบบทำงานแบบเดิม / เลข PR รันต่อเนื่อง / อนุมัติแล้วไม่เขียน Purchase History /
// รับของต้องเลือก PR / รับบางส่วน-รับเกิน / คืนรายการแล้วหักยอดรับกลับ / ปิดยอดค้าง / ยกเลิกทั้งใบ
const { backend, grabFn, grabVar, grabScalar, assert } = require('./_export-extract');

function makeSheet(rows) {
  const data = rows || [];
  return {
    rows: data,
    getLastRow: function() { return data.length; },
    getLastColumn: function() { return data.reduce(function(m, r) { return Math.max(m, r.length); }, 0); },
    appendRow: function(row) { data.push(row.slice()); },
    getDataRange: function() {
      return { getValues: function() { return data.map(function(r) { return r.slice(); }); } };
    },
    getRange: function(row, col, numRows, numCols) {
      return {
        getValues: function() {
          const out = [];
          for (let r = 0; r < (numRows || 1); r += 1) {
            const src = data[row - 1 + r] || [];
            const line = [];
            for (let c = 0; c < (numCols || 1); c += 1) line.push(src[col - 1 + c] === undefined ? '' : src[col - 1 + c]);
            out.push(line);
          }
          return out;
        },
        setValues: function(values) {
          for (let r = 0; r < (numRows || 1); r += 1) {
            data[row - 1 + r] = data[row - 1 + r] || [];
            for (let c = 0; c < (numCols || 1); c += 1) data[row - 1 + r][col - 1 + c] = values[r][c];
          }
        },
        setValue: function(v) { data[row - 1] = data[row - 1] || []; data[row - 1][col - 1] = v; }
      };
    }
  };
}

function makeBackend() {
  const clock = { now: Date.parse('2026-10-07T03:00:00Z') }; // 10:00 เวลาไทย
  const sheets = {};
  const props = {};
  const phWrites = [];
  const ADMIN_PERMS = { view: true, pr_create: true, pr_view_own: true, pr_view_all: true, pr_approve: true, manage_users: true };
  const users = {
    admin: { username: 'admin', role: 'admin', isActive: true, line: '', permissions: ADMIN_PERMS },
    tech: { username: 'tech', role: 'user', isActive: true, line: 'STAR', permissions: { view: true, pr_create: true } }
  };
  const tokens = { 'tok-admin': 'admin', 'tok-tech': 'tech' };

  class FakeDate extends Date {
    constructor(...args) { if (args.length) super(...args); else super(clock.now); }
    static now() { return clock.now; }
    static [Symbol.hasInstance](v) { return Object.prototype.toString.call(v) === '[object Date]'; }
  }
  const p = function(n) { return ('0' + n).slice(-2); };
  const Utilities = {
    formatDate: function(date, tz, fmt) {
      const t = new Date(new Date(date).getTime() + 7 * 3600000);
      const y = String(t.getUTCFullYear()), mo = p(t.getUTCMonth() + 1), d = p(t.getUTCDate());
      const hh = p(t.getUTCHours()), mi = p(t.getUTCMinutes()), ss = p(t.getUTCSeconds());
      if (fmt === 'yyMM') return y.slice(2) + mo;
      if (fmt === 'yyyyMMdd-HHmmss') return y + mo + d + '-' + hh + mi + ss;
      return y + '-' + mo + '-' + d + ' ' + hh + ':' + mi + ':' + ss;
    }
  };
  function getSessionUser(payload) {
    const name = tokens[payload && payload.authToken];
    if (!name) throw new Error('กรุณาเข้าสู่ระบบ');
    return { user: { username: name } };
  }
  function requirePermission(payload, perm) {
    const u = users[getSessionUser(payload).user.username];
    if (!u.permissions[perm]) throw new Error('ไม่มีสิทธิ์ใช้งานฟังก์ชันนี้ (' + perm + ')');
    return u;
  }
  function requireAdminUser(payload) {
    const u = requirePermission(payload, 'manage_users');
    if (u.role !== 'admin') throw new Error('เฉพาะ Admin');
    return u;
  }
  const env = {
    SPARE_APP_CONFIG: { prHeaderSheetName: 'PRHeaders', prLinesSheetName: 'PRLines', prAuditSheetName: 'PRAudit', grLogSheetName: 'GRLog' },
    SpreadsheetApp: { getActiveSpreadsheet: function() { return {}; } },
    getOrCreateSheet: function(ss, name) { sheets[name] = sheets[name] || makeSheet([]); return sheets[name]; },
    LockService: { getScriptLock: function() { return { waitLock: function() {}, releaseLock: function() {} }; } },
    PropertiesService: { getScriptProperties: function() { return { getProperty: function(k) { return props[k] || null; }, setProperty: function(k, v) { props[k] = v; } }; } },
    Utilities: Utilities,
    Logger: { log: function() {} },
    getSessionUser: getSessionUser,
    requirePermission: requirePermission,
    requireAdminUser: requireAdminUser,
    findUserByUsername: function(name) { return users[name] || null; },
    upsertPurchaseHistoryRecord: function(rec) { phWrites.push(rec); }
  };

  const fnNames = [
    'prStr', 'getPrSheetWithHeaders', 'getPrHeaderSheet', 'getPrLinesSheet', 'getPrAuditSheet', 'prIndexMap', 'appendPrAudit',
    'parsePrLinesInput', 'prUserCanAccessLine', 'prHeaderRowToCard', 'findPrHeaderRow', 'createPRUnlocked', 'approvePRUnlocked',
    'getPRForApproval', 'toBoolean', 'normalizeRole', 'normalizeLogTimestamp',
    'normalizePurchaseHistoryModel', 'isMeaningfulPurchaseHistoryModel', 'normalizePurchaseHistoryName',
    'getGrLogSheet', 'getPrGrCutoverAt', 'isPrGrSystemEnabled', 'prNowText', 'getPrGrSystemStatus', 'startPrGrSystem',
    'generatePrRunningId', 'prLineIsClosed', 'prLineOutstanding', 'prLineStatusFromQty', 'computePrHeaderStatus', 'prIsTracked',
    'prLineRowToObject', 'refreshPrHeaderStatus', 'findPrLineRow', 'prLineMatchesReceivedPart', 'resolvePrReceiptForTransaction',
    'buildGrId', 'grLogRef', 'postPrGoodsReceipt', 'reversePrGoodsReceiptForLog', 'markPROrderedUnlocked', 'cancelPRUnlocked',
    'closePrLineUnlocked', 'listPRs', 'getPRDetail', 'grRowToObject', 'listGrRows', 'getGrLog', 'listOpenPrLines', 'getOpenPrLines'
  ];
  const src = [
    grabVar('PR_HEADER_HEADERS', backend), grabVar('PR_LINE_HEADERS', backend), grabVar('PR_AUDIT_HEADERS', backend),
    grabVar('GR_LOG_HEADERS', backend), grabVar('PR_RECEIVABLE_STATUSES', backend), grabScalar('PR_GR_CUTOVER_PROP', backend)
  ].concat(fnNames.map(function(n) { return grabFn(n, backend, 0); }))
    .concat(['return {' + fnNames.map(function(n) { return n + ': ' + n; }).join(', ') + '};']).join('\n');
  const envKeys = Object.keys(env);
  const api = new Function(envKeys.concat(['Date']).join(','), src).apply(null, envKeys.map(function(k) { return env[k]; }).concat([FakeDate]));
  api.sheets = sheets;
  api.props = props;
  api.phWrites = phWrites;
  api.clock = clock;
  return api;
}

const LINES = JSON.stringify([
  { part_no: '12', part_name: 'Cutter blade', model: 'CB-100', brand: 'ACME', unit: 'PCS', qty: 10, unit_price: 50 },
  { part_no: '15', part_name: 'Spring', model: '-', unit: 'PCS', qty: 4, unit_price: 20 }
]);
function receivePayload(extra) {
  return Object.assign({ authToken: 'tok-tech', partNo: '12', partName: 'Cutter blade', model: 'CB-100', brand: 'ACME', process: 'STAR', unit: 'PCS', by: 'tech' }, extra || {});
}
function receive(api, extra, qty, ts) {
  const payload = receivePayload(extra);
  const receipt = api.resolvePrReceiptForTransaction(payload);
  if (receipt.mode === 'LEGACY') return { receipt: receipt };
  return { receipt: receipt, gr: api.postPrGoodsReceipt(receipt, payload, qty, ts, 'Main List Stock') };
}
function lineOf(api, prId, lineNo) {
  return api.getPRDetail({ authToken: 'tok-admin', pr_id: prId }).lines.filter(function(l) { return l.line_no === lineNo; })[0];
}

// ---- ก่อนเปิดระบบ: ทุกอย่างเหมือนเดิม (deploy ได้ปลอดภัย) ----
(function legacyBeforeCutover() {
  const api = makeBackend();
  assert.strictEqual(api.isPrGrSystemEnabled(), false);
  assert.deepStrictEqual(api.resolvePrReceiptForTransaction(receivePayload()), { mode: 'LEGACY' }, 'ยังไม่เปิดระบบ = รับเข้าแบบเดิม ไม่บังคับ PR');
  const res = api.createPRUnlocked({ authToken: 'tok-tech', pr_id: 'PR-CLIENT-1', line: 'STAR', lines_json: LINES });
  assert.strictEqual(res.pr_id, 'PR-CLIENT-1', 'ระบบเดิมใช้เลข PR ที่หน้าเว็บส่งมา');
  api.approvePRUnlocked({ authToken: 'tok-admin', pr_id: 'PR-CLIENT-1' });
  assert.strictEqual(api.phWrites.length, 2, 'ระบบเดิม: อนุมัติแล้วเขียน Purchase History ทุกบรรทัด');
  assert.strictEqual(api.listOpenPrLines({}).length, 0, 'ยังไม่เปิดระบบ = ไม่มีรายการติดตาม');
})();

// ---- เปิดระบบ → PR ใหม่ถูกติดตาม, PR เก่าไม่ถูกนับรอของเข้า ----
(function fullLifecycle() {
  const api = makeBackend();
  api.createPRUnlocked({ authToken: 'tok-tech', pr_id: 'PR-OLD-1', line: 'STAR', lines_json: LINES });
  api.approvePRUnlocked({ authToken: 'tok-admin', pr_id: 'PR-OLD-1' });

  assert.throws(function() { api.startPrGrSystem({ authToken: 'tok-tech' }); }, /manage_users/, 'เริ่มระบบได้เฉพาะ Admin');
  const started = api.startPrGrSystem({ authToken: 'tok-admin' });
  assert.strictEqual(started.cutover_at, '2026-10-07 10:00:00');
  assert.strictEqual(api.startPrGrSystem({ authToken: 'tok-admin' }).already_started, true, 'กดซ้ำไม่รีเซ็ตวันเริ่ม');
  assert.strictEqual(api.listOpenPrLines({}).length, 0, 'PR เก่าที่อนุมัติไว้ไม่ถูกนับเป็นรอของเข้า (เริ่มจากศูนย์)');

  // เลข PR ออกจาก server แบบรันต่อเนื่อง ไม่ใช้เลขที่หน้าเว็บสุ่มมา
  const pr1 = api.createPRUnlocked({ authToken: 'tok-tech', pr_id: 'PR-RANDOM-999', line: 'STAR', lines_json: LINES, assign_to: 'admin' });
  const pr2 = api.createPRUnlocked({ authToken: 'tok-tech', line: 'STAR', lines_json: LINES });
  assert.strictEqual(pr1.pr_id, 'PR-2610-001');
  assert.strictEqual(pr2.pr_id, 'PR-2610-002');

  // รอนุมัติ = ยังรับของไม่ได้
  assert.throws(function() { receive(api, { prId: 'PR-2610-001', prLineNo: 1 }, 1, '2026-10-07 10:05:00'); }, /ยังรับของไม่ได้ \(สถานะ: PENDING\)/);

  const phBefore = api.phWrites.length;
  api.approvePRUnlocked({ authToken: 'tok-admin', pr_id: 'PR-2610-001', lines_json: JSON.stringify([{ line_no: 2, qty_approved: 3 }]) });
  assert.strictEqual(api.phWrites.length, phBefore, 'ใบระบบใหม่: อนุมัติแล้วไม่เขียน Purchase History');
  let open = api.listOpenPrLines({});
  assert.strictEqual(open.length, 2, 'อนุมัติแล้วเข้ารายการค้างรับ');
  assert.deepStrictEqual(open.map(function(l) { return l.qty_outstanding; }), [10, 3], 'ยอดค้างใช้ qty_approved');

  // dropdown "รับตาม PR" กรองเฉพาะอะไหล่ตัวที่กำลังรับ
  const forCutter = api.getOpenPrLines({ authToken: 'tok-tech', part_name: 'Cutter blade', model: 'CB-100' }).lines;
  assert.strictEqual(forCutter.length, 1);
  assert.strictEqual(forCutter[0].line_no, 1);

  // ไม่เลือก PR: คนทั่วไปรับไม่ได้ / Admin ต้องใส่เหตุผล
  assert.throws(function() { receive(api, {}, 1, '2026-10-07 10:10:00'); }, /กรุณาเลือก PR/);
  assert.throws(function() { receive(api, { noPrReason: 'ของแถม' }, 1, '2026-10-07 10:10:00'); }, /เฉพาะ Admin/);
  const noPr = receive(api, { authToken: 'tok-admin', noPrReason: 'ของแถมจากซัพ' }, 2, '2026-10-07 10:10:00');
  assert.strictEqual(noPr.gr.mode, 'NO_PR');
  const noPrRow = api.listGrRows({}).filter(function(g) { return g.gr_id === noPr.gr.gr_id; })[0];
  assert.strictEqual(noPrRow.amount, 0, 'รับโดยไม่มี PR ไม่คิดเป็นค่าใช้จ่าย');
  assert.strictEqual(noPrRow.no_pr_reason, 'ของแถมจากซัพ');

  // อะไหล่ไม่ตรงกับบรรทัด PR = ล้ม
  assert.throws(function() { receive(api, { prId: 'PR-2610-001', prLineNo: 2 }, 1, '2026-10-07 10:12:00'); }, /ไม่ตรงกับรายการใน PR/);

  // สั่งซื้อแล้ว (ใส่ PO) → รับบางส่วน
  assert.throws(function() { api.markPROrderedUnlocked({ authToken: 'tok-tech', pr_id: 'PR-2610-001' }); }, /pr_view_all/);
  api.markPROrderedUnlocked({ authToken: 'tok-admin', pr_id: 'PR-2610-001', po_no: 'PO-77', vendor: 'ACME Co.' });
  const r1 = receive(api, { prId: 'PR-2610-001', prLineNo: 1 }, 4, '2026-10-08 09:00:00');
  assert.strictEqual(r1.gr.gr_id, 'GR-2610-0002', 'เลข GR รันต่อจากใบก่อนหน้า');
  assert.strictEqual(r1.gr.qty_outstanding, 6);
  assert.strictEqual(r1.gr.line_status, 'PARTIAL');
  assert.strictEqual(r1.gr.pr_status, 'PARTIAL');
  assert.strictEqual(api.listGrRows({ pr_id: 'PR-2610-001' })[0].amount, 200, 'มูลค่ารับ = จำนวน × ราคาใน PR');

  // รับเกินได้ — บันทึกส่วนเกินไว้
  const r2 = receive(api, { prId: 'PR-2610-001', prLineNo: 1 }, 8, '2026-10-09 09:00:00');
  assert.strictEqual(r2.gr.qty_received, 12);
  assert.strictEqual(r2.gr.over_qty, 2, 'รับเกิน 2 ชิ้นถูกบันทึกเป็น over_qty');
  assert.strictEqual(r2.gr.line_status, 'RECEIVED');
  assert.strictEqual(r2.gr.pr_status, 'PARTIAL', 'ยังเหลือบรรทัด 2 ค้างอยู่');
  assert.throws(function() { receive(api, { prId: 'PR-2610-001', prLineNo: 1 }, 1, '2026-10-09 09:30:00'); }, /รับครบแล้ว/);

  // คืนรายการรับเข้า (จากหน้า Log) → ยกเลิก GR + หักยอดรับกลับ
  const reversed = api.reversePrGoodsReceiptForLog('2026-10-09 09:00:00', 'Cutter blade', 'admin');
  assert.strictEqual(reversed.gr_id, r2.gr.gr_id);
  assert.strictEqual(reversed.qty_received, 4);
  assert.strictEqual(lineOf(api, 'PR-2610-001', 1).line_status, 'PARTIAL');
  assert.strictEqual(api.listGrRows({ pr_id: 'PR-2610-001' }).length, 1, 'GR ที่ถูกคืนไม่ถูกนับ');
  assert.strictEqual(api.listGrRows({ pr_id: 'PR-2610-001', include_reversed: true }).length, 2, 'แต่ยังเก็บไว้เป็นหลักฐาน');
  assert.strictEqual(api.reversePrGoodsReceiptForLog('2026-10-09 09:00:00', 'Cutter blade', 'admin'), null, 'คืนซ้ำไม่หักซ้ำ');

  // ปิดยอดค้าง (ของไม่มาแล้ว) ต้องมีเหตุผล → ทั้งใบปิด
  assert.throws(function() { api.closePrLineUnlocked({ authToken: 'tok-admin', pr_id: 'PR-2610-001', line_no: 1 }); }, /เหตุผล/);
  api.closePrLineUnlocked({ authToken: 'tok-admin', pr_id: 'PR-2610-001', line_no: 1, reason: 'ซัพไม่มีของ' });
  const springPayload = { prId: 'PR-2610-001', prLineNo: 2, partNo: '15', partName: 'Spring', model: '-' };
  const r3 = receive(api, springPayload, 3, '2026-10-10 09:00:00');
  assert.strictEqual(r3.gr.pr_status, 'CLOSED', 'มีบรรทัดถูกปิดยอดค้าง = ใบปิด (ไม่ใช่รับครบ)');
  const detail = api.getPRDetail({ authToken: 'tok-admin', pr_id: 'PR-2610-001' });
  assert.strictEqual(detail.header.status, 'CLOSED');
  assert.strictEqual(detail.header.po_no, 'PO-77');
  assert.strictEqual(detail.receipts.length, 2);
  const actions = detail.timeline.map(function(t) { return t.action; });
  ['CREATE', 'APPROVE', 'ORDER', 'RECEIVE', 'RECEIVE_REVERSED', 'CLOSE_LINE', 'STATUS_CLOSED'].forEach(function(a) {
    assert(actions.indexOf(a) > -1, 'ไทม์ไลน์ต้องมี ' + a);
  });
  assert.strictEqual(api.listOpenPrLines({}).length, 0, 'PR-2610-002 ยังไม่อนุมัติ จึงไม่มีอะไรค้างรับ');

  // ยกเลิกทั้งใบ: ใบที่ยังไม่มีของเข้าได้ / ใบที่มีของเข้าแล้วไม่ได้
  assert.throws(function() { api.cancelPRUnlocked({ authToken: 'tok-tech', pr_id: 'PR-2610-002' }); }, /เหตุผล/);
  assert.strictEqual(api.cancelPRUnlocked({ authToken: 'tok-tech', pr_id: 'PR-2610-002', reason: 'สั่งซ้ำ' }).pr_status, 'CANCELLED');
  assert.throws(function() { api.cancelPRUnlocked({ authToken: 'tok-admin', pr_id: 'PR-2610-001', reason: 'x' }); }, /ยกเลิกทั้งใบไม่ได้/);

  // PR เก่าก่อนเปิดระบบยังอยู่ครบ ดูย้อนหลังได้ แต่รับของตามใบเก่าไม่ได้
  const all = api.listPRs({ authToken: 'tok-admin' }).prs.map(function(c) { return c.pr_id + ':' + c.gr_tracked; });
  assert.deepStrictEqual(all.sort(), ['PR-2610-001:true', 'PR-2610-002:true', 'PR-OLD-1:false']);
  assert.throws(function() { receive(api, { prId: 'PR-OLD-1', prLineNo: 1 }, 1, '2026-10-10 10:00:00'); }, /ใบระบบเก่า/);
})();

// ---- สถานะใบคำนวณจากบรรทัด (pure) ----
(function headerStatusRules() {
  const api = makeBackend();
  const c = api.computePrHeaderStatus;
  assert.strictEqual(c('ORDERED', [{ qty_approved: 5, qty_received: 0 }]), 'ORDERED');
  assert.strictEqual(c('ORDERED', [{ qty_approved: 5, qty_received: 2 }]), 'PARTIAL');
  assert.strictEqual(c('PARTIAL', [{ qty_approved: 5, qty_received: 7 }]), 'RECEIVED', 'รับเกินก็นับว่ารับครบ');
  assert.strictEqual(c('PARTIAL', [{ qty_approved: 5, qty_received: 0 }], true), 'ORDERED', 'คืนรายการจนยอดเป็นศูนย์ ถอยกลับ ORDERED');
  assert.strictEqual(c('PARTIAL', [{ qty_approved: 5, qty_received: 0 }], false), 'APPROVED');
  assert.strictEqual(c('PENDING', [{ qty_approved: 5, qty_received: 5 }]), 'PENDING', 'ยังไม่อนุมัติไม่ถูกเปลี่ยนจากยอดรับ');
  assert.strictEqual(c('APPROVED', [{ qty_approved: 5, qty_received: 5 }, { qty_approved: 0, qty_received: 0, line_status: 'CANCELLED' }]), 'RECEIVED');
})();

// ---- จุดเชื่อมใน processTransaction / returnLogEntry / Inbox / routing ----
(function wiring() {
  const txn = backend.slice(backend.indexOf('function processTransactionUnlocked(payload)'), backend.indexOf('function resolveReadSheetName('));
  const resolveAt = txn.indexOf('resolvePrReceiptForTransaction(payload)');
  const stockWriteAt = txn.indexOf('mainSheet.getRange(sheetRowNumber, stockCol + 1).setValue(stockAfter);');
  const postAt = txn.indexOf('postPrGoodsReceipt(prReceipt, payload, signedQty, txnTimestamp, resolvedSheetName)');
  assert(resolveAt > -1 && stockWriteAt > -1 && postAt > -1);
  assert(resolveAt < stockWriteAt, 'ตรวจ PR ก่อนแตะสต็อก — PR ผิดต้องล้มทั้งรายการ');
  assert(postAt > txn.indexOf('historySheet.appendRow(['), 'บันทึก GR หลังเขียน Log (ใช้ txnTimestamp เดียวกันเป็น log_ref)');
  assert(txn.includes('goodsReceipt: goodsReceipt'), 'ส่งผลใบรับของกลับให้หน้าเว็บ');

  const ret = backend.slice(backend.indexOf('function returnLogEntryUnlocked(payload)'), backend.indexOf('function processTransaction(payload)'));
  assert(ret.includes('signedQty > 0 ? reversePrGoodsReceiptForLog(originalTs, originalName, user.username) : null'), 'คืนรายการรับเข้า → ยกเลิก GR');

  const get = backend.slice(backend.indexOf('function parseTransactionPayloadFromGet(e)'), backend.indexOf('function getLogRows('));
  ['prId: e.parameter.prId', 'prLineNo: e.parameter.prLineNo', 'noPrReason: e.parameter.noPrReason'].forEach(function(f) {
    assert(get.includes(f), 'GET transact ต้องส่ง ' + f);
  });

  const inbox = backend.slice(backend.indexOf('function getInbox(payload)'), backend.indexOf('function listBudgetBypasses('));
  assert(/if \(isPrGrSystemEnabled\(\)\) \{[\s\S]*?listOpenPrLines\(\{\}\)/.test(inbox), 'Inbox ของรอเข้าอ่านจาก PR ค้างรับเมื่อเปิดระบบ');

  ['getPrGrSystemStatus', 'startPrGrSystem', 'markPROrdered', 'cancelPR', 'closePrLine', 'listPRs', 'getPRDetail', 'getGrLog', 'getOpenPrLines'].forEach(function(action) {
    assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(e.parameter), e);"), 'doGet routes ' + action);
    assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(body), e);"), 'doPost routes ' + action);
  });
})();

console.log('PR → GR tracking checks passed');

// ---- Google Sheets แปลงสตริงวันที่เป็น Date เอง — ต้องยังคำนวณวันรอ/กรองเดือนได้ ----
(function dateCellsFromSheets() {
  const api = makeBackend();
  api.startPrGrSystem({ authToken: 'tok-admin' });
  const pr = api.createPRUnlocked({ authToken: 'tok-tech', line: 'STAR', lines_json: LINES });
  api.approvePRUnlocked({ authToken: 'tok-admin', pr_id: pr.pr_id });
  const rows = api.sheets.PRHeaders.rows;
  const idx = rows[0].indexOf('approved_at');
  const cIdx = rows[0].indexOf('created_at');
  rows[1][idx] = new Date('2026-10-02T03:00:00Z'); // 2026-10-02 10:00 เวลาไทย เป็น Date จริง
  rows[1][cIdx] = new Date('2026-10-01T03:00:00Z');
  const open = api.listOpenPrLines({});
  assert.strictEqual(open[0].days_waiting, 5, 'วันรอนับจากวันอนุมัติ แม้เซลล์เป็น Date');
  assert.strictEqual(api.listPRs({ authToken: 'tok-admin', month: '2026-10' }).prs.length, 1, 'กรองเดือนได้แม้ created_at เป็น Date');
})();

console.log('PR → GR date-cell checks passed');
