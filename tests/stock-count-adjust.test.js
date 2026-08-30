const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const backendLf = backend.replace(/\r\n/g, '\n');

// ── Frontend: scDoSubmit ต้องประกาศ sessionCopy ก่อนใช้ ────────────────────────
// var hoisting: ถ้าประกาศ var sessionCopy ทีหลังจุดที่อ่านค่า จะได้ undefined แล้ว throw
// ตั้งแต่บรรทัดแรกของ .then() → คิว pending_approval ไม่ถูกเขียนลง localStorage เลย
// ผลคือ Engineer ไม่มีอะไรให้กดอนุมัติ และ Stock ไม่มีวันถูกปรับ
const submitBlock = htmlLf.slice(
  htmlLf.indexOf('function scDoSubmit()'),
  htmlLf.indexOf('function scGetFilteredItems()')
);
assert(submitBlock, 'ต้องหาบล็อก scDoSubmit เจอ');
assert((submitBlock.match(/var sessionCopy = JSON\.parse\(JSON\.stringify\(scSession\)\);/g) || []).length === 1,
  'ต้องประกาศ sessionCopy ครั้งเดียวใน scDoSubmit');
const declPos = submitBlock.indexOf('var sessionCopy =');
const firstUsePos = submitBlock.indexOf('sessionCopy.month');
assert(declPos > -1 && firstUsePos > -1);
assert(declPos < firstUsePos,
  'ต้องประกาศ sessionCopy ก่อนบรรทัดแรกที่ใช้ ไม่งั้น var hoisting จะทำให้เป็น undefined');
// ต้อง copy ก่อนล้าง scSession ไม่งั้นได้ null
assert(declPos < submitBlock.indexOf('scSession = null;'),
  'ต้อง copy scSession ก่อนตั้งเป็น null');
// รายการที่ส่งแล้วต้องเข้าคิวรออนุมัติ ไม่ปรับ Stock ทันที
assert(submitBlock.includes("status: 'pending_approval'"));
assert(!/adjustStockFromCount/.test(submitBlock),
  'scDoSubmit ห้ามปรับ Stock เอง — ต้องรอ Engineer อนุมัติ');

// ── Backend: ปรับยอดจากการนับ ห้ามสร้าง Purchase History ───────────────────────
// นับได้เกินระบบ → ลงเป็น Input ซึ่ง processTransaction จะ sync purchase history ให้
// (ทำให้ยอดค่าใช้จ่ายเดือนนั้นบวมเกินจริง) และ stamp ป้าย "ของใหม่" ทั้งที่ไม่ได้ซื้อของเข้ามา
const adjustBlock = backendLf.slice(
  backendLf.indexOf('function adjustStockFromCount(payload)'),
  backendLf.indexOf('// SMART AUTOMATION + AI FEATURES')
);
assert(adjustBlock, 'ต้องหาบล็อก adjustStockFromCount เจอ');
assert(adjustBlock.includes('skipPurchaseHistory: true'),
  'txnPayload ของ adjustStockFromCount ต้องมี skipPurchaseHistory: true');
assert(adjustBlock.indexOf('skipPurchaseHistory: true') < adjustBlock.indexOf('processTransaction(txnPayload)'),
  'ต้องใส่ flag ก่อนเรียก processTransaction');
// flag นี้ต้องยังกันทั้ง purchase history และป้ายของใหม่อยู่
assert(backend.includes('if (signedQty > 0 && !payload.skipPurchaseHistory) {'));
assert((backend.match(/if \(signedQty > 0 && !payload\.skipPurchaseHistory\) \{/g) || []).length === 2,
  'ต้องกันทั้ง stampLastReceivedAt และ syncPurchaseHistoryOnReceive');

// ปรับยอดเป็นการลง Input/Output ด้วยส่วนต่าง ไม่ใช่เขียนทับตัวเลข — ต้องมี audit trail
assert(adjustBlock.includes('var variance = Number(item.counted) - Number(item.system_qty);'));
assert(adjustBlock.includes("type: variance > 0 ? 'Input' : 'Output'"));
assert(adjustBlock.includes("reason: 'Stock Adjustment'"));

console.log('stock-count-adjust: OK');
