// บั๊กจริง: รับเข้า MXH10-30Z (cyl, compact slide table) 1 ชิ้น แต่ Log ขึ้น Input +1 สามแถว
// เวลา 10:39 เดียวกัน สต็อก 0→1→2→3 — server บันทึกแล้วแต่หน้าเว็บขึ้น "ไม่สำเร็จ" → กดซ้ำ / ส่งซ้ำผ่านช่องทางสำรอง
// เทสต์นี้รันโค้ดกันซ้ำจริงทั้งฝั่ง Backend และหน้าเว็บ
const { grabFn, assert } = require('./_export-extract');
const fs = require('fs');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

// ---------- Backend: processTransaction ตอบผลเดิมเมื่อ id ซ้ำ ไม่แตะสต็อกอีก ----------
const beFns = ['processTransaction', 'normalizeClientTxnId', 'transactionDedupeFingerprint', 'readTransactionDedupe', 'rememberTransactionDedupe']
  .map(function(n) { return grabFn(n, backend, 0); }).join('\n');
function buildBackend() {
  const store = {};
  const state = { stock: 0, logRows: 0, locks: 0 };
  const api = new Function('LockService', 'CacheService', 'Logger', 'processTransactionUnlocked',
    'var TRANSACTION_DEDUPE_TTL_SECONDS = 21600;\n' + beFns + '\nreturn processTransaction;')(
    { getScriptLock: function() { return { waitLock: function() { state.locks++; }, releaseLock: function() { state.locks--; } }; } },
    { getScriptCache: function() { return { get: function(k) { return store[k] || null; }, put: function(k, v) { store[k] = v; } }; } },
    { log: function() {} },
    function(p) {
      const before = state.stock;
      state.stock += Number(p.qty);
      state.logRows++;
      return { status: 'success', stockBefore: before, stockAfter: state.stock };
    });
  return { run: api, state: state };
}
const be = buildBackend();
const p = { type: 'Input', sheetName: 'Lug&Screw', partNo: '12', partName: 'cyl, compact slide table', model: 'MXH10-30Z', qty: '1', clientTxnId: 'TXN-abc-123' };
const r1 = be.run(p);
const r2 = be.run(Object.assign({}, p));
const r3 = be.run(Object.assign({}, p, { qty: 1, sheetName: undefined, sheet: 'Lug&Screw' })); // ส่งซ้ำอีกช่องทาง (POST ใช้ key sheet)
assert.strictEqual(be.state.stock, 1, 'กดซ้ำ 3 ครั้งด้วย id เดิม สต็อกต้องบวกแค่ครั้งเดียว');
assert.strictEqual(be.state.logRows, 1, 'Log ต้องมีแถวเดียว');
assert(!r1.duplicate && r2.duplicate && r3.duplicate, 'ครั้งที่ 2-3 ตอบว่าเป็นรายการซ้ำ');
assert.strictEqual(r2.stockAfter, 1, 'ตอบผลเดิม (Stock 0 → 1)');
assert.strictEqual(be.state.locks, 0, 'ปล่อย lock ทุกครั้ง');
// id ใหม่ = รายการใหม่ (รับของจริงอีกรอบ) ต้องบันทึกได้ตามปกติ
be.run(Object.assign({}, p, { clientTxnId: 'TXN-abc-456' }));
assert.strictEqual(be.state.stock, 2);
// id เดิมแต่เนื้อรายการต่าง (จำนวนต่าง) ห้ามตอบผลเก่าผิดตัว
be.run(Object.assign({}, p, { clientTxnId: 'TXN-abc-456', qty: '5' }));
assert.strictEqual(be.state.stock, 7);
// ไม่มี id (เวอร์ชันหน้าเว็บเก่า / คืนรายการภายใน) = ทำงานเหมือนเดิม
be.run(Object.assign({}, p, { clientTxnId: '' }));
be.run(Object.assign({}, p, { clientTxnId: '' }));
assert.strictEqual(be.state.stock, 9);
// id แปลกปลอมไม่ถูกใช้เป็นคีย์
const be2 = buildBackend();
be2.run(Object.assign({}, p, { clientTxnId: 'bad id with spaces' }));
be2.run(Object.assign({}, p, { clientTxnId: 'bad id with spaces' }));
assert.strictEqual(be2.state.stock, 2);
// GET ต้องส่ง clientTxnId ต่อให้ processTransaction
assert(grabFn('parseTransactionPayloadFromGet', backend, 0).includes('clientTxnId: e.parameter.clientTxnId'));

// ---------- หน้าเว็บ: รายการเดิมได้ id เดิมจนกว่าจะสำเร็จ ----------
const feFns = ['txnIntentFingerprint', 'getTxnClientId', 'clearTxnClientId', 'callTransactionViaJsonp'].map(function(n) { return grabFn(n); }).join('\n');
async function runFrontend() {
  const sent = [];
  let reply = function() { return Promise.reject(new Error('Load failed')); };
  const toasts = [];
  const fe = new Function('sendTransactionRequest', 'showQuickToast',
    'var TXN_SUBMIT_TICKET_TTL_MS = 30 * 60 * 1000; var txnSubmitTickets = {};\n' + feFns + '\nreturn callTransactionViaJsonp;')(
    function(payload, id) { sent.push(id); return reply(); },
    function(msg) { toasts.push(msg); });
  const recv = { type: 'Input', sheetName: 'Lug&Screw', partNo: '12', partName: 'cyl, compact slide table', model: 'MXH10-30Z', qty: 1 };
  await fe(Object.assign({}, recv)).catch(function() {});
  await fe(Object.assign({}, recv)).catch(function() {});
  assert.strictEqual(sent[0], sent[1], 'ขึ้น "ไม่สำเร็จ" แล้วกดซ้ำ ต้องส่ง id เดิม');
  assert(/^TXN-[a-z0-9]+-[a-z0-9]+$/.test(sent[0]), 'id อยู่ในรูปแบบที่ backend รับ');
  reply = function() { return Promise.resolve({ status: 'success', duplicate: true, stockBefore: 0, stockAfter: 1 }); };
  await fe(Object.assign({}, recv));
  assert.strictEqual(sent[2], sent[0], 'ครั้งที่ 3 ยังเป็น id เดิม');
  assert(toasts.some(function(t) { return t.indexOf('บันทึกไปแล้ว') > -1; }), 'บอกผู้ใช้ว่ารายการนี้บันทึกไปแล้ว');
  reply = function() { return Promise.resolve({ status: 'success' }); };
  await fe(Object.assign({}, recv));
  assert.notStrictEqual(sent[3], sent[0], 'สำเร็จแล้ว รายการถัดไปต้องได้ id ใหม่ (รับของจริงอีกรอบได้)');
  await fe(Object.assign({}, recv, { qty: 2 }));
  assert.notStrictEqual(sent[4], sent[3], 'จำนวนต่าง = รายการใหม่');
}

// sendTransactionRequest ต้องส่ง id ทั้ง POST และ query ของช่องทางสำรอง
const send = grabFn('sendTransactionRequest');
assert(send.includes("clientTxnId: clientTxnId || ''") && send.includes("'clientTxnId=' + encodeURIComponent(requestPayload.clientTxnId)"));

runFrontend().then(function() {
  console.log('transaction-duplicate-guard: OK');
}).catch(function(err) { console.error(err); process.exit(1); });
