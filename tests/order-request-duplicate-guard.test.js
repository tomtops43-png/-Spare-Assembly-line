const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// บั๊กจริง: ผู้ใช้กด "ส่งคำขอซื้อ" → ของเข้าชีตแล้ว แต่เน็ตหลุดตอนรับคำตอบ
// เว็บขึ้น "ส่งคำขอไม่สำเร็จ" ทั้งที่สำเร็จ ผู้ใช้เลยกดส่งใหม่ → ได้คำขอซ้ำ 2 ใบ
// (เคสในภาพ: "ประแจเลื่อน" 2 ใบ ห่างกัน 47 วินาที ข้อมูลเหมือนกันเป๊ะ)

// ---- Backend: มีคอลัมน์กุญแจกันซ้ำ ----
assert(backend.includes("'unit', 'unit_price', 'currency', 'client_uid'"), 'ORDER_REQUEST_HEADERS gains client_uid');

// ---- Backend: createOrderRequest ล็อก + เช็คซ้ำก่อน append ----
assert(/function createOrderRequest\(payload\) \{[\s\S]{0,400}var lock = LockService\.getScriptLock\(\);\s*lock\.waitLock\(30000\);/.test(backend), 'createOrderRequest uses LockService like approvePR');
assert(backend.includes('function findOrderRequestIdByClientUid(sheet, idx, clientUid)'), 'lookup by client_uid exists');
assert(backend.includes('var duplicateOf = findOrderRequestIdByClientUid(sheet, idx, clientUid);'), 'checks for an existing row before appending');
assert(/if \(duplicateOf\) \{[\s\S]{0,300}return \{ status: 'success', request_id: duplicateOf, duplicate: true/.test(backend), 'duplicate submit returns the original request as success, not a new row');
assert(backend.includes('lock.releaseLock();'), 'lock is always released');

// ---- Backend: เขียนแถวตามหัวคอลัมน์จริง ไม่ใช่ลำดับตายตัว ----
assert(backend.includes('var row = headers.map(function(h) {'), 'row is built against the live sheet headers');
assert(backend.includes('client_uid: clientUid'), 'client_uid is persisted on the row');
assert(backend.includes('syncPurchaseHistoryForRequest(toRequestObject(headers, row)'), 'purchase history sync uses the same header order');

// ---- Frontend: uid ผูกกับเนื้อคำขอ ส่งซ้ำใบเดิมได้ uid เดิม ----
assert(html.includes('function orderRequestFingerprint(src)'), 'fingerprint helper exists');
assert(html.includes('function getRoSubmitUid(payload)'), 'submit ticket helper exists');
assert(html.includes('if (roSubmitTicket.uid && roSubmitTicket.fingerprint === fp) return roSubmitTicket.uid;'), 'retrying the same request reuses the uid');
assert(html.includes('payload.client_uid = getRoSubmitUid(payload);'), 'submit attaches client_uid');
assert(html.includes('function clearRoSubmitTicket()'), 'ticket is cleared after a confirmed success');
assert(/finishOrderRequestSubmit\(res\) \{[\s\S]{0,900}clearRoSubmitTicket\(\);/.test(html), 'success path clears the ticket so the next request gets a fresh uid');

// ---- Frontend: ห้ามขึ้น "ไม่สำเร็จ" ก่อนเช็คกับ server ----
assert(html.includes('function findSubmittedOrderRequest(clientUid, fingerprint)'), 'verification helper exists');
assert(/\.catch\(function\(err\) \{[\s\S]{0,600}return findSubmittedOrderRequest\(submitUid, submitFingerprint\)\.then\(function\(landed\) \{/.test(html), 'failure path verifies with the server before reporting failure');
assert(html.includes('if (landed) { finishOrderRequestSubmit({ verified: true }); return; }'), 'a request that actually landed is reported as success');
assert(html.includes("String(rows[i].client_uid || '').trim() === clientUid"), 'primary match is by client_uid');
assert(html.includes('if (roKnownRequestIdsLoaded) {'), 'heuristic fallback only runs when the known-id snapshot exists');
assert(html.includes('ระบบกันคำขอซ้ำให้แล้ว'), 'error toast tells the user it is safe to retry');

// ---- Frontend: จำ request_id ที่เห็นแล้ว เพื่อแยกใบที่เพิ่งโผล่ ----
assert(html.includes('var roKnownRequestIds = {};'), 'known-id map declared');
assert(html.includes('var roKnownRequestIdsLoaded = false;'), 'known-id loaded flag declared');
assert(/roKnownRequestIds = \{\};[\s\S]{0,300}roKnownRequestIdsLoaded = true;/.test(html), 'loadOrderRequests refreshes the known-id snapshot');

console.log('order-request-duplicate-guard: all assertions passed');
