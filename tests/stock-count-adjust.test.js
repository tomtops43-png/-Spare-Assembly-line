const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const backendLf = backend.replace(/\r\n/g, '\n');

function slice(src, from, to, label) {
  const a = src.indexOf(from);
  const b = src.indexOf(to);
  assert(a > -1 && b > a, 'ต้องหาบล็อก ' + label + ' เจอ');
  return src.slice(a, b);
}

// ── Frontend: scDoSubmit ต้องประกาศ sessionCopy ก่อนใช้ ────────────────────────
// var hoisting: ถ้าประกาศ var sessionCopy ทีหลังจุดที่อ่านค่า จะได้ undefined แล้ว throw
// ตั้งแต่บรรทัดแรกของ .then() → ส่งผลนับแล้วขึ้น "บันทึกไม่สำเร็จ" ทั้งที่ลงชีทไปแล้ว
const submitBlock = slice(htmlLf, 'function scDoSubmit()', 'function scGetFilteredItems()', 'scDoSubmit');
assert((submitBlock.match(/var sessionCopy = JSON\.parse\(JSON\.stringify\(scSession\)\);/g) || []).length === 1,
  'ต้องประกาศ sessionCopy ครั้งเดียวใน scDoSubmit');
const declPos = submitBlock.indexOf('var sessionCopy =');
const firstUsePos = submitBlock.indexOf('sessionCopy.month');
assert(declPos > -1 && firstUsePos > -1);
assert(declPos < firstUsePos,
  'ต้องประกาศ sessionCopy ก่อนบรรทัดแรกที่ใช้ ไม่งั้น var hoisting จะทำให้เป็น undefined');
assert(declPos < submitBlock.indexOf('scSession = null;'),
  'ต้อง copy scSession ก่อนตั้งเป็น null');
assert(!/adjustStockFromCount/.test(submitBlock),
  'scDoSubmit ห้ามยิง adjustStockFromCount ตรงๆ — ต้องผ่าน approveStockCount ที่ gate สิทธิ์ไว้');

// ── Frontend: คนที่อนุมัติได้อยู่แล้วไม่ต้องส่งให้ใครตรวจ ────────────────────────
// Admin/Engineer กดส่งผลแล้วปรับ Stock ทันที · คนอื่นยังเข้าคิว pending_approval ตามเดิม
assert(submitBlock.includes('if (!scCanApprove()) return { sessionId: sessionId, adjusted: 0, failed: 0, autoApproved: false };'),
  'ไม่มีสิทธิ์อนุมัติต้องเข้าคิวตามเดิม');
assert(submitBlock.includes("action: 'approveStockCount'"),
  'มีสิทธิ์อนุมัติต้องอนุมัติของตัวเองต่อทันที');
assert(submitBlock.includes('loadPartsData({ skipCache: true })'),
  'ปรับ Stock แล้วต้องโหลดข้อมูลอะไหล่ใหม่ ไม่งั้นตารางค้างยอดเก่า');
// ปรับ Stock ไม่ผ่านต้องไม่ทำให้ผลนับหาย — ตกไปเข้าคิวรออนุมัติแทน
assert(/\.catch\(function\(err\) \{[\s\S]{0,400}autoApproved: false/.test(submitBlock),
  'ปรับ Stock ล้มต้อง fallback เป็นเข้าคิว ไม่ใช่ทิ้งผลนับ');
const submitSessionBlock = slice(htmlLf, 'function scSubmitSession()', 'function scDoSubmit()', 'scSubmitSession');
assert(submitSessionBlock.includes('scCanApprove() && diffCount > 0'),
  'กดแล้วยอดสต็อกเปลี่ยนทันที ต้องถามยืนยันก่อน');
assert(htmlLf.includes('function scSubmitLabelText()') && htmlLf.includes('function scSyncSubmitLabel()'),
  'ป้ายปุ่มต้องเปลี่ยนตามสิทธิ์ ไม่ใช่ค้างว่า "ส่งให้ Engineer ตรวจสอบ" ตลอด');

// ── Backend: ปรับยอดจากการนับ ห้ามสร้าง Purchase History ───────────────────────
// นับได้เกินระบบ → ลงเป็น Input ซึ่ง processTransaction จะ sync purchase history ให้
// (ทำให้ยอดค่าใช้จ่ายเดือนนั้นบวมเกินจริง) และ stamp ป้าย "ของใหม่" ทั้งที่ไม่ได้ซื้อของเข้ามา
const adjustBlock = slice(backendLf, 'function adjustStockFromCount(payload)',
  '// SMART AUTOMATION + AI FEATURES', 'adjustStockFromCount');
assert(adjustBlock.includes('skipPurchaseHistory: true'),
  'txnPayload ของ adjustStockFromCount ต้องมี skipPurchaseHistory: true');
assert(adjustBlock.indexOf('skipPurchaseHistory: true') < adjustBlock.indexOf('processTransaction(txnPayload)'),
  'ต้องใส่ flag ก่อนเรียก processTransaction');
assert((backend.match(/if \(signedQty > 0 && !payload\.skipPurchaseHistory\) \{/g) || []).length === 2,
  'flag นี้ต้องยังกันทั้ง stampLastReceivedAt และ syncPurchaseHistoryOnReceive');

// ปรับยอดเป็นการลง Input/Output ด้วยส่วนต่าง ไม่ใช่เขียนทับตัวเลข — ต้องมี audit trail
assert(adjustBlock.includes('var variance = Number(item.counted) - Number(item.system_qty);'));
assert(adjustBlock.includes("type: variance > 0 ? 'Input' : 'Output'"));
assert(adjustBlock.includes("reason: 'Stock Adjustment'"));

// endpoint ที่แก้ยอดสต็อกได้ตรงๆ ต้องไม่เปิดให้ช่างที่มีแค่สิทธิ์กรอกยอดนับ
assert(adjustBlock.includes("requirePermission({ authToken: payload.authToken }, 'manage_items')"),
  'adjustStockFromCount ต้อง gate ด้วย manage_items');
assert(!/requirePermission\(\{ authToken: payload\.authToken \}, 'view_logs'\)/.test(adjustBlock),
  "adjustStockFromCount ต้องไม่ gate ด้วย 'view_logs' (กว้างเกินไป)");

// ── Backend: คิวอนุมัติต้องอยู่บนชีท ไม่ใช่ localStorage ─────────────────────────
assert(backend.includes('function approveStockCount(payload)'));
assert((backend.match(/action === 'approveStockCount'/g) || []).length === 2,
  'ต้อง dispatch approveStockCount ทั้ง doGet และ doPost');
assert((backend.match(/action === 'getStockCountHistory'/g) || []).length === 2,
  'ต้อง dispatch getStockCountHistory ทั้ง doGet และ doPost (หน้าเว็บเรียกผ่าน POST)');

const approveBlock = slice(backendLf, 'function approveStockCount(payload)',
  'function adjustStockFromCount(payload)', 'approveStockCount');
assert(approveBlock.includes("requirePermission({ authToken: payload.authToken }, 'manage_items')"),
  'approveStockCount ต้อง gate ด้วย manage_items');
// items ต้องอ่านจากชีท ไม่รับจาก client — กันแก้ตัวเลขที่นับได้ระหว่างทางก่อนอนุมัติ
assert(approveBlock.includes('data[rowIndex][idx.items_json]'),
  'approveStockCount ต้องอ่าน items จากชีท');
assert(!/payload\.(items|diff_items)/.test(approveBlock),
  'approveStockCount ห้ามรับ items/diff_items จาก client');
// กันกดอนุมัติซ้ำ
assert(approveBlock.includes('isStockCountPending(currentStatus)'));
assert(approveBlock.includes('ถูกดำเนินการไปแล้ว'));
// ห้ามจับ script lock ครอบ adjustStockFromCount — ข้างในเรียก processTransaction ที่จับ lock เอง
// (โปรเจกต์นี้ใช้ *Unlocked variant กันซ้อน เช่น returnLogEntryUnlocked → processTransactionUnlocked)
assert(!/LockService/.test(approveBlock),
  'approveStockCount ห้ามจับ script lock ซ้อนกับ processTransaction');

// สถานะ + คอลัมน์ผู้อนุมัติ
assert(backend.includes("'approved_by','approved_at','adjusted_count'"),
  'STOCK_COUNT_HEADERS ต้องมีคอลัมน์ผู้อนุมัติ');
assert(backend.includes("var STOCK_COUNT_PENDING_STATUSES = ['', 'submitted', 'pending_approval'];"),
  "แถวเก่าที่สถานะเป็น 'submitted' ต้องยังนับเป็นรออนุมัติ");
assert(backendLf.includes("    'pending_approval',"),
  'saveStockCountResult ต้องเขียนสถานะ pending_approval');
// ชีทเก่ามี 13 คอลัมน์ — ต้องเติมหัวที่ขาด "ต่อท้าย" เท่านั้น ห้ามเขียนทับหัวเดิม
assert(backend.includes('function ensureStockCountHeaders(sheet)'));
const ensureBlock = slice(backendLf, 'function ensureStockCountHeaders(sheet)',
  'function stockCountIndexMap(headers)', 'ensureStockCountHeaders');
assert(ensureBlock.includes('sheet.getRange(1, lastCol + 1, 1, missing.length)'),
  'ต้องเติมคอลัมน์ต่อท้าย ไม่ใช่เขียนทับทั้งแถวหัว');

// ── Frontend: ต้องเลิกใช้ localStorage เป็นคิวอนุมัติ ────────────────────────────
// (sc_history ยังถูกอ่านอยู่ แต่เพื่อ "ยกขึ้นเซิร์ฟเวอร์" อย่างเดียว ไม่ใช่ใช้เป็นคิว)
const pendingBlock = slice(htmlLf, 'function scRenderPending()', 'function scDecide(btn)', 'scRenderPending');
assert(!/localStorage/.test(pendingBlock),
  'คิวรออนุมัติห้ามอ่าน localStorage — Engineer ต้องอนุมัติจากเครื่องไหนก็ได้');
const historyBlock = slice(htmlLf, 'function scRenderHistory()', 'function scRenderPending()', 'scRenderHistory');
assert(!/localStorage/.test(historyBlock), 'ประวัติห้ามอ่าน localStorage');
assert(html.includes("action: 'getStockCountHistory'"), 'ประวัติ/คิว ต้องดึงจากเซิร์ฟเวอร์');
assert(html.includes("action: 'approveStockCount'"), 'ปุ่มอนุมัติต้องยิงไป approveStockCount');
const decideBlock = slice(htmlLf, 'function scDecide(btn)', '// ── Stock Count event bindings', 'scDecide');
assert(!/\b(items|diff_items)\s*:/.test(decideBlock),
  'ฝั่งหน้าเว็บห้ามส่ง items ไปตอนอนุมัติ — Backend อ่านจากชีทเอง');
assert(decideBlock.includes("decision: decision"));
assert(decideBlock.includes('loadPartsData({ skipCache: true })'),
  'ปรับ Stock แล้วต้องโหลดข้อมูลอะไหล่ใหม่ ไม่งั้นตารางค้างยอดเก่า');
// เปิดหน้ามาต้องเห็นคิวเลย ไม่ต้องกดปุ่มดูประวัติก่อน
assert(htmlLf.includes('if (scCanApprove()) return scRefresh();'));

// ── เก็บประวัติเก่าไว้เป็นหลักฐาน ───────────────────────────────────────────────
// เวอร์ชันแรกของฟีเจอร์ (cb174ff) บันทึกผลนับลง localStorage อย่างเดียว ไม่เคยยิงขึ้นเซิร์ฟเวอร์
// ผลนับรอบนั้นจึงมีอยู่ที่เดียวคือเบราว์เซอร์ของช่าง ต้องยกขึ้นชีทก่อนหาย
assert(backend.includes('function importLegacyStockCount(payload)'));
assert((backend.match(/action === 'importLegacyStockCount'/g) || []).length === 2,
  'ต้อง dispatch importLegacyStockCount ทั้ง doGet และ doPost');
const legacyBlock = slice(backendLf, 'function importLegacyStockCount(payload)',
  'function approveStockCount(payload)', 'importLegacyStockCount');
// ของเก่าเก็บเป็นหลักฐานอย่างเดียว ห้ามหลุดเข้าคิวอนุมัติแล้วไปปรับ Stock ด้วยยอดที่เก่าเป็นเดือน
assert(/archived_legacy/.test(legacyBlock) && /archived_approved/.test(legacyBlock) && /archived_rejected/.test(legacyBlock),
  'สถานะของ record ที่นำเข้าต้องขึ้นต้นด้วย archived_');
const pendingStatuses = backend.match(/STOCK_COUNT_PENDING_STATUSES = \[[^\]]*\]/)[0];
assert(!/archived/.test(pendingStatuses),
  'สถานะ archived_* ห้ามถูกนับเป็นรออนุมัติ — เก็บเป็นหลักฐานอย่างเดียว');
assert(/appendRow/.test(legacyBlock), 'ต้อง append แถวใหม่ ไม่ใช่เขียนทับของเดิม');
assert(legacyBlock.includes("imported: false"), 'ต้องกันนำเข้าซ้ำด้วย session_id');

// ห้ามลบ localStorage ทิ้งหลังยกขึ้นเซิร์ฟเวอร์ — เก็บไว้เป็น backup ในเครื่อง
assert(!/removeItem\(\s*(SC_LEGACY_KEY|'sc_history')/.test(html),
  'ห้ามลบ sc_history ทิ้ง — เก็บไว้เป็นหลักฐานสำรองในเครื่อง');
assert(html.includes("action: 'importLegacyStockCount'"), 'หน้าเว็บต้องยกประวัติเก่าขึ้นเซิร์ฟเวอร์');
const migrateBlock = slice(htmlLf, 'function scMigrateLegacyHistory()', 'function scLineLabel(line)', 'scMigrateLegacyHistory');
assert(migrateBlock.includes('legacy_status: h.status'),
  'ต้องยกผลตัดสินเดิม (อนุมัติ/ส่งคืน) ขึ้นไปด้วย ไม่งั้นประวัติเพี้ยน');
assert(migrateBlock.includes('items: JSON.stringify(h.items || [])'),
  'ต้องยกรายการที่นับได้ขึ้นไปด้วย ไม่งั้นเก็บไว้เป็นหลักฐานไม่ได้');
// ยกไม่สำเร็จต้องไม่ตั้ง flag — ครั้งหน้าจะได้ลองใหม่ ไม่ใช่ปล่อยหลักฐานค้างในเครื่องเดียว
const doneIdx = migrateBlock.indexOf('localStorage.setItem(SC_LEGACY_DONE_KEY');
assert(doneIdx > -1 && doneIdx < migrateBlock.indexOf('.catch(function(err)'),
  'ต้องตั้ง flag ใน then เท่านั้น ห้ามตั้งใน catch');

// ── ชีทเป็น append-only: ห้ามมีอะไรลบแถวประวัติเช็คสต็อกทิ้ง ─────────────────────
const stockCountArea = slice(backendLf, 'var STOCK_COUNT_SHEET_NAME',
  '// SMART AUTOMATION + AI FEATURES', 'stock count backend');
assert(!/deleteRow|clearContent|clear\(\)/.test(stockCountArea),
  'โค้ดฝั่งเช็คสต็อกห้ามลบ/ล้างแถวในชีท — เก็บเป็นหลักฐานทั้งหมด');
// อนุมัติแล้วต้องไม่ล้าง items_json ทิ้ง — ยอดที่ช่างนับได้คือตัวหลักฐาน
assert(!/idx\.items_json \+ 1/.test(approveBlock),
  'approveStockCount ห้ามเขียนทับคอลัมน์ items_json');

// ── Backend: ต้องใช้สเปรดชีตที่สคริปต์ผูกอยู่ ────────────────────────────────────
// getOrCreateStockCountSheet เคยเรียก SpreadsheetApp.openById(SPREADSHEET_ID) โดยที่
// SPREADSHEET_ID ไม่เคยถูกประกาศไว้ที่ไหนเลย → ทุก action ของหน้าเช็คสต็อกพังหมด
// ("SPREADSHEET_ID is not defined") ฟังก์ชันอื่นในไฟล์ใช้ getActiveSpreadsheet() กันทั้งหมด
assert(!/SPREADSHEET_ID/.test(backend),
  'ห้ามอ้างตัวแปร SPREADSHEET_ID ที่ไม่มีใครประกาศ');
const sheetGetterBlock = slice(backendLf, 'function getOrCreateStockCountSheet()',
  'function ensureStockCountHeaders(sheet)', 'getOrCreateStockCountSheet');
assert(sheetGetterBlock.includes('SpreadsheetApp.getActiveSpreadsheet()'),
  'ต้องใช้ getActiveSpreadsheet() เหมือนฟังก์ชันอื่นในไฟล์');
assert(!/SpreadsheetApp\.openById\(/.test(sheetGetterBlock));

// ── Frontend: ยอดที่นับไปแล้วห้ามหายเพราะรีเฟรช/บันทึกไม่ผ่าน ──────────────────
assert(htmlLf.includes("var SC_DRAFT_KEY = 'sc_draft_session';"));
assert(htmlLf.includes('function scRestoreDraft()') && htmlLf.includes('function scSaveDraft()'));
// เซฟทุกทางที่ยอดเปลี่ยน: พิมพ์เอง / ใส่ 0 ทั้งหมด / คัดลอกจากระบบ / เปิด Session
assert((htmlLf.match(/scSaveDraft\(\);/g) || []).length >= 3, 'ต้องเซฟ draft ทุกจุดที่ยอดเปลี่ยน');
assert(htmlLf.includes('scSaveDraftSoon();'), 'ช่องกรอกต้องเซฟแบบ debounce');
assert(htmlLf.includes('scRestoreDraft();'), 'เปิดหน้ามาต้องถามกู้ draft');
// ล้าง draft ได้เฉพาะตอนบันทึกขึ้นชีทสำเร็จ หรือกดยกเลิก Session เอง — ห้ามล้างตอน error
assert(submitBlock.indexOf('scClearDraft();') > -1 &&
  submitBlock.indexOf('scClearDraft();') < submitBlock.indexOf('scSession = null;'),
  'ต้องล้าง draft ในทางสำเร็จ ก่อนปิด session');
const submitCatch = submitBlock.slice(submitBlock.indexOf("showQuickToast('บันทึกไม่สำเร็จ"));
assert(!/scClearDraft/.test(submitCatch),
  'บันทึกไม่สำเร็จห้ามล้าง draft — ยอดที่นับมาทั้งวันต้องยังอยู่');

// ── ปรับยอดต้องลงชีทต้นทางของอะไหล่ ไม่ใช่ชีท default ─────────────────────────
// processTransaction ใช้ resolveReadSheetName(payload.sheetName) ถ้าไม่ส่งมาจะ fallback ไป
// SPARE_APP_CONFIG.readSheetName ('Main List Stock') — นับชีท Coil Winding แล้วไปแก้ยอดที่
// Main List Stock: ถ้าหาไม่เจอก็ error เงียบๆ (ยอดเหมือนเดิม) ถ้าเจอชื่อซ้ำก็แก้ผิดชีท
assert(adjustBlock.includes("sheetName: String(item.sheet || '')"),
  'txnPayload ต้องระบุชีทต้นทาง ไม่งั้นปรับผิดชีท');
assert(adjustBlock.indexOf('sheetName') < adjustBlock.indexOf('processTransaction(txnPayload)'));
// ไม่รู้ชีท = ต้องล้มรายการนั้นให้เห็น ไม่ใช่เดาแล้วไปแก้ชีทอื่น
assert(adjustBlock.includes('ไม่รู้ว่าอะไหล่นี้อยู่ชีทไหน'),
  'ไม่รู้ชีทต้นทางต้อง throw ไม่ใช่ fallback ไปชีท default');
assert(adjustBlock.indexOf('ไม่รู้ว่าอะไหล่นี้อยู่ชีทไหน') < adjustBlock.indexOf('var txnPayload'),
  'ต้องเช็คก่อนสร้าง txnPayload');

// ชีทต้นทางต้องถูกส่งต่อครบทั้งสาย: เปิด session → บันทึกขึ้นชีท → อนุมัติ → ปรับยอด
const startBlock = slice(htmlLf, 'function scStartSession()', 'function scSubmitLabelText()', 'scStartSession');
assert(/sheet: p\.__sourceSheet/.test(startBlock), 'ตอนเปิด session ต้องเก็บชีทต้นทางของแต่ละรายการ');
// โหลดแยกทีละ sheet แล้วรวมพูล — ต้องแปะ __sourceSheet ให้แต่ละแถวตอนโหลด ก่อนรวม ไม่งั้นรวมแล้ว
// แยกไม่ออกว่าแถวไหนมาจากชีทไหน (backend ไม่ได้ใส่ฟิลด์ sheet มาให้ในแถวเอง) ทำให้ p.__sourceSheet
// ว่างเปล่าทุกแถว → ส่งผลนับได้ปกติ แต่ตอนอนุมัติแล้วปรับ Stock ล้มทุกรายการเพราะไม่รู้ชีทต้นทาง
assert(/\.map\(function\(sheetName\)[\s\S]{0,700}__sourceSheet\s*=\s*sheetName/.test(startBlock),
  'ตอนโหลดแต่ละ sheet ต้องแปะ __sourceSheet = sheetName ให้ทุกแถวก่อนรวมพูล ไม่งั้น sheet ต้นทางหายหมด');
assert(/category: p\.category/.test(startBlock), 'ต้องเก็บ category ด้วย ไม่งั้น Log ลงเป็น General หมด');
assert(/sheet: it\.sheet/.test(submitBlock), 'ตอนส่งผลต้องแนบชีทต้นทางขึ้นชีทด้วย');
assert(/id: it\.id/.test(submitBlock) && /brand: it\.brand/.test(submitBlock),
  'ต้องแนบ id/brand ด้วย ไม่งั้นจับคู่อะไหล่ไม่แม่น');
assert(/sheet: it\.sheet \|\| ''/.test(approveBlock), 'approveStockCount ต้องส่งชีทต้นทางต่อ');

console.log('stock-count-adjust: OK');
