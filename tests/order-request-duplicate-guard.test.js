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

// ---- Frontend: ตั๋วต้องรอดข้ามการรีเฟรช ----
// รูรั่วเดิม: uid อยู่แค่ในหน่วยความจำ ผู้ใช้เจอ error → ปิดแอป/รีเฟรช → กดส่งใหม่
// รอบสองได้ uid ใหม่ server เลยแยกไม่ออกว่าเป็นใบเดิม → ได้คำขอซ้ำอยู่ดี
assert(html.includes("var RO_SUBMIT_TICKET_KEY = 'spare_ro_submit_ticket';"), 'ticket has a localStorage key');
assert(html.includes('function readRoSubmitTicket()'), 'ticket can be restored from the device');
assert(html.includes('localStorage.setItem(RO_SUBMIT_TICKET_KEY, JSON.stringify(roSubmitTicket));'), 'ticket is persisted when issued');
assert(/saveRoSubmitTicket\(\);\s*return roSubmitTicket\.uid;/.test(html), 'a freshly issued uid is saved before it is handed out');
assert(/var stored = readRoSubmitTicket\(\);\s*if \(stored && stored\.fingerprint === fp\) \{ roSubmitTicket = stored; return roSubmitTicket\.uid; \}/.test(html),
  'after a reload the stored ticket is reused when the request content matches');

// ตั๋วค้างต้องหมดอายุ ไม่งั้นคำขอที่ตั้งใจขอซ้ำของเดิมทีหลังจะโดนกันไปด้วย
assert(html.includes('var RO_SUBMIT_TICKET_TTL_MS ='), 'ticket has a TTL');
assert(html.includes('Date.now() - Number(t.savedAt) > RO_SUBMIT_TICKET_TTL_MS) return null;'), 'expired tickets are ignored');

// เครื่องที่ใช้ร่วมกัน: คนถัดไปที่ขอของเหมือนกันเป๊ะต้องได้ใบของตัวเอง ไม่ไปชนใบของคนก่อนหน้า
assert(html.includes('function roSubmitTicketOwner()'), 'ticket is scoped to the signed-in user');
assert(html.includes("if (String(t.owner || '') !== roSubmitTicketOwner()) return null;"), 'another user on the same device gets a fresh uid');

// อ่าน/เขียน localStorage ไม่ได้ (โหมดส่วนตัว, storage เต็ม) ต้องไม่ทำให้ส่งคำขอพัง
assert(/function readRoSubmitTicket\(\) \{[\s\S]{0,1200}\} catch \(err\) \{\s*return null;/.test(html), 'unreadable storage falls back to an in-memory ticket');
assert(/function saveRoSubmitTicket\(\) \{[\s\S]{0,600}\} catch \(err\) \{/.test(html), 'unwritable storage does not break submitting');
assert(/function clearRoSubmitTicket\(\) \{[\s\S]{0,400}localStorage\.removeItem\(RO_SUBMIT_TICKET_KEY\)/.test(html), 'clearing the ticket also clears the device copy');

console.log('order-request-duplicate-guard: all assertions passed');
