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

// ── มือถือ: การ์ดแทนตาราง ─────────────────────────────────────────────────────
// ตารางเดิมเป็น min-w-[700px] ในกรอบเลื่อนแนวนอน บนมือถือต้องรูดซ้ายขวาตลอดและคอลัมน์
// รุ่น/ตำแหน่ง/เหตุผล ถูกซ่อนหมด — ช่างที่เดินนับหน้างานใช้ไม่ได้จริง
assert(html.includes('id="scCardList"'), 'ต้องมีลิสต์การ์ดสำหรับมือถือ');
assert(/id="scCardList"[^>]*class="[^"]*md:hidden/.test(html), 'การ์ดต้องโชว์เฉพาะจอเล็ก');
assert(/<div class="hidden md:block overflow-x-auto">/.test(html), 'ตารางเดิมต้องถูกซ่อนบนจอเล็ก');
assert(html.includes('function scRenderCards()'), 'ต้องมีตัว render การ์ด');
// ทั้งสองมุมมองแสดงรายการชุดเดียวกัน ต้องวาดพร้อมกันเสมอ ไม่งั้นสลับจอแล้วข้อมูลไม่ตรง
const renderBlock = slice(htmlLf, 'function scRenderTable()', 'function scStartSession()', 'scRenderTable');
assert(renderBlock.includes('scRenderCards();'), 'scRenderTable ต้องวาดการ์ดด้วยเสมอ');

// รูปต้องกดดูเต็มจอได้ — ใช้ lightbox ตัวเดิมของหน้า Stock ที่ pinch-zoom ได้อยู่แล้ว
const cardBlock = slice(htmlLf, 'function scRenderCards()', 'function scRenderTable()', 'scRenderCards');
assert(/data-sc-img="/.test(cardBlock), 'การ์ดต้องมีปุ่มรูปที่พกลิงก์รูปไว้');
assert(/openImageViewer\(imgBtn\.getAttribute\('data-sc-img'\)\)/.test(htmlLf),
  'กดรูปต้องเปิด lightbox เต็มจอ');
// ปุ่มบนมือถือต้องกดด้วยนิ้วโป้งได้ — h-11 = 44px ตามเกณฑ์ touch target
assert(/data-sc-step="-1"[^>]*class="h-11 w-11/.test(cardBlock) || /class="h-11 w-11[^"]*"[^>]*data-sc-step/.test(cardBlock) ||
  /h-11 w-11[\s\S]{0,200}data-sc-step/.test(cardBlock) || /data-sc-step[\s\S]{0,200}h-11 w-11/.test(cardBlock),
  'ปุ่ม −/+ ต้องสูงอย่างน้อย 44px (h-11)');
assert(/inputmode="numeric"/.test(cardBlock), 'ช่องกรอกยอดบนมือถือต้องเรียกแป้นตัวเลข');

// ปุ่มส่งผลต้องติดขอบล่างจอ — นับ 75 รายการแล้วเลื่อนกลับขึ้นไปหาปุ่มบนสุดไม่ไหว
assert(html.includes('id="scMobileBar"'), 'ต้องมีแถบส่งผลติดขอบล่างจอ');
assert(html.includes('id="scSubmitBtnMobile"'), 'แถบล่างต้องมีปุ่มส่งผล');
assert(/scSubmitBtnMobile[\s\S]{0,200}addEventListener\('click', scSubmitSession\)/.test(htmlLf),
  'ปุ่มส่งบนแถบล่างต้องทำงานเหมือนปุ่มบนสุด');
// แถบล่างต้องโผล่/หายตามสถานะ session จากจุดเดียว ไม่ใช่ไล่ toggle ทีละที่แล้วหลุด
const statsBlock = slice(htmlLf, 'function scUpdateStats()', 'function scRenderRoundBanner', 'scUpdateStats');
assert(/scMobileBar[\s\S]{0,160}classList\.toggle\('hidden', !scSession\)/.test(statsBlock),
  'แถบล่างต้องผูกกับ !scSession ใน scUpdateStats');

// ── สิทธิ์: Admin เท่านั้นที่ปรับยอดได้ ─────────────────────────────────────────
const canApprove = slice(htmlLf, 'function scCanApprove()', 'var SC_DRAFT_KEY', 'scCanApprove');
assert(/return isAdminUser\(\);/.test(canApprove), 'scCanApprove ต้องเหลือแค่ isAdminUser()');
assert(!/manage_items/.test(canApprove),
  "ห้ามใช้ manage_items — role 'leader' ได้สิทธิ์นั้นติดมาโดยปริยาย จะทำให้ Leader ปรับยอดได้");
assert(html.includes('ส่งผลให้ Admin ตรวจสอบ'), 'ป้ายปุ่มของคนที่ไม่ใช่ Admin ต้องบอกว่าส่งให้ Admin');
// Backend ต้อง gate ซ้ำอีกชั้น ไม่งั้นยิง API ตรงก็ยังปรับยอดได้ทั้งที่หน้าเว็บซ่อนปุ่มแล้ว
assert(backend.includes('function requireStockCountAdmin(payload)'), 'ต้องมี gate ฝั่ง Backend');
['adjustStockFromCount', 'approveStockCount', 'approveStockCountGroup', 'getStockCountComparison'].forEach(function (fn) {
  const at = backendLf.indexOf('function ' + fn + '(payload)');
  assert(at > -1, 'ต้องหาฟังก์ชัน ' + fn + ' เจอ');
  // gate ต้องอยู่ต้นฟังก์ชัน ก่อนจะไปแตะข้อมูลอะไร — ดูช่วงหัวฟังก์ชันก็พอ
  const head = backendLf.slice(at, at + 700);
  assert(/requireStockCountAdmin\(/.test(head), fn + ' ต้อง gate ด้วย requireStockCountAdmin ที่ต้นฟังก์ชัน');
});

// ── นับซ้ำทาน: รอบแรกปิดยอดระบบ ───────────────────────────────────────────────
assert(htmlLf.includes('blind: !isAdminUser()'),
  'ช่าง/Leader ต้องนับแบบปิดยอดระบบเสมอ — Admin เป็นคนกระทบยอดจึงเห็นได้');
assert(html.includes('function scIsBlind()'), 'ต้องมีตัวเช็คโหมดปิดยอด');
// สีเขียว/แดงบนการ์ดบอกยอดระบบทางอ้อม (แดง = ไม่ตรงระบบ) ตอนปิดยอดต้องไม่ใช้เกณฑ์นี้
assert(html.includes('function scGetDisplayStatus(item)'), 'ต้องแยกสถานะที่เอาไปแสดงผลออกจากสถานะจริง');
assert(cardBlock.includes('scGetDisplayStatus(it)'), 'การ์ดต้องระบายสีตาม display status');
assert(!/var s = scGetStatusClass\(it\);/.test(cardBlock),
  'การ์ดห้ามระบายสีตาม scGetStatusClass ตรง ๆ — จะเฉลยยอดระบบตอนนับรอบแรก');
// สถิติ/ตัวกรองก็เฉลยได้เหมือนกัน ต้องปิดตอนนับรอบแรก
assert(/if \(diffTile\) diffTile\.classList\.toggle\('hidden', blind\)/.test(statsBlock),
  'ตอนปิดยอดต้องซ่อนตัวเลข "ไม่ตรง" ในแถบสถิติ');
assert(/diffPanel\.classList\.toggle\('hidden', blind \|\| /.test(statsBlock),
  'ตอนปิดยอดต้องซ่อนกล่องสรุปรายการที่ไม่ตรงกับระบบ');
assert(/if \(key === 'match' \|\| key === 'diff'\) chip\.classList\.toggle\('hidden', hideCompare\)/.test(htmlLf),
  'ตอนปิดยอดและยังไม่มียอดรอบก่อน ต้องซ่อนชิปกรอง ตรงกัน/ไม่ตรง');
// คอลัมน์ "ในระบบ"/"ส่วนต่าง" ของตารางคอมก็ต้องปิดตาม ไม่ใช่ปิดแค่ฝั่งการ์ด
assert(renderBlock.includes(".sc-sys-col") && /classList\.toggle\('hidden', scIsBlind\(\)\)/.test(renderBlock),
  'คอลัมน์ยอดระบบของตารางต้องถูกซ่อนตอนนับแบบปิดยอด');

// ── ปิดยอดต้องปิดทุกทางออก ไม่ใช่แค่ในลิสต์ที่นับ ──────────────────────────────
// ซ่อนแค่ในตาราง/การ์ดไม่พอ ถ้ายังกดปุ่มอื่นแล้วเห็นยอดระบบได้อยู่ ฟีเจอร์นับปิดยอดก็ไร้ผล
const printBlock = slice(htmlLf, 'function scPrintCountSheet()', 'function scExportCountCsv()', 'scPrintCountSheet');
assert(/var blindPrint = scIsBlind\(\);/.test(printBlock), 'ใบปริ้นต้องรู้ว่ากำลังนับแบบปิดยอด');
assert(/blindPrint \? '—' : esc\(it\.systemQty\)/.test(printBlock),
  'ใบปริ้นห้ามพิมพ์ยอดระบบลงไปตอนปิดยอด — ซ่อนด้วย CSS ไม่พอ เปิด view-source ก็เห็น');
const csvBlock = slice(htmlLf, 'function scExportCountCsv()', 'function scExportPdf', 'scExportCountCsv');
assert(/var blindCsv = scIsBlind\(\);/.test(csvBlock) && /blindCsv \? '' : it\.systemQty/.test(csvBlock),
  'CSV ที่ Export ตอนปิดยอด ต้องไม่มียอดระบบติดไป');
// คอลัมน์ต้องยังอยู่ครบ เพราะ scApplyImportedCsv หาคอลัมน์จากชื่อหัว ตัดออกแล้วนำเข้ากลับไม่ได้
assert(/'ยอดในระบบ'/.test(csvBlock), 'หัวคอลัมน์ต้องคงไว้ ไม่งั้นนำเข้า CSV กลับไม่ได้');

const submitBlock = slice(htmlLf, 'function scDoSubmit()', 'function scGetFilteredItems()', 'scDoSubmit');
// ต้องอ่านโหมดปิดยอดก่อนล้าง scSession ไม่งั้น scIsBlind() เป็น false เสมอตอนวาดกล่องสรุป
assert(submitBlock.indexOf('var submitBlind = scIsBlind();') > -1 &&
  submitBlock.indexOf('var submitBlind = scIsBlind();') < submitBlock.indexOf('scSession = null;'),
  'ต้องอ่านโหมดปิดยอดก่อนล้าง scSession');
assert(/submitBlind[\s\S]{0,900}📝 นับแล้ว/.test(submitBlock),
  'กล่องสรุปของคนที่นับปิดยอด ต้องบอกแค่ "นับแล้วกี่รายการ" ไม่ใช่ ตรงกัน/ไม่ตรง/ความแม่น');
assert(/\(submitBlind \? '' :\s*\n?\s*'<button id="_scResultPdf"/.test(submitBlock),
  'ห้ามให้คนที่นับปิดยอดโหลด PDF รายงาน — ในนั้นมีคอลัมน์ยอดระบบ + ส่วนต่าง');
const historyBlock = slice(htmlLf, 'function scRenderHistory()', 'function scRenderPending()', 'scRenderHistory');
assert(/var showAccuracy = scCanApprove\(\);/.test(historyBlock),
  'ประวัติต้องโชว์ความแม่น/ตรงกี่รายการ เฉพาะ Admin');
// ส่งซ้ำได้ (ทับผลของตัวเอง) ถ้าเห็นว่า "ตรง 52/74" ย้อนหลัง ก็ย้อนไปแก้ให้ตรงระบบได้
assert(/showAccuracy[\s\S]{0,320}รายการ<\/p>/.test(historyBlock),
  'คนที่ไม่ใช่ Admin ต้องเห็นแค่จำนวนรายการ ไม่ใช่เปอร์เซ็นต์ความแม่น');

// draft ที่เซฟไว้ก่อนมีฟีเจอร์นี้ไม่มีฟิลด์ blind (undefined = เห็นยอดระบบ) ต้องคำนวณใหม่เสมอ
const restoreBlock = slice(htmlLf, 'function scRestoreDraft()', 'function scResumeSessionUi()', 'scRestoreDraft');
assert(/scSession\.blind = !isAdminUser\(\);/.test(restoreBlock),
  'กู้ draft ต้องคำนวณโหมดปิดยอดใหม่จากผู้ใช้ปัจจุบัน');

// ── รอบ 2 ขึ้นไป: เห็นยอดของรอบก่อน ───────────────────────────────────────────
assert(html.includes('function scItemKey(name, model, sheet)'), 'หน้าเว็บต้องมีคีย์จับคู่อะไหล่');
assert(backend.includes('function stockCountItemKey(name, model, sheet)'), 'Backend ต้องมีคีย์จับคู่อะไหล่');
assert(html.includes('prevCount:'), 'item ต้องพกยอดของรอบก่อนมาด้วย');
assert(/scSession\.roundNo/.test(htmlLf) && html.includes('function scRenderRoundBanner()'),
  'ต้องบอกผู้ใช้ว่ากำลังนับรอบที่เท่าไหร่');
// รอบก่อนหน้าต้องถูกดึงมาก่อนสร้าง session ไม่งั้น item จะไม่มี prevCount ติดไปเลย
const startBlock = slice(htmlLf, 'function scStartSession()', 'function scSubmitLabelText()', 'scStartSession');
assert(startBlock.indexOf("action: 'getStockCountGroupState'") > -1,
  'ต้องถามสถานะรอบนับจากเซิร์ฟเวอร์ก่อนเปิด session');
assert(startBlock.indexOf('groupStatePromise') < startBlock.indexOf('scSession = {'),
  'ต้องได้ผลรอบนับมาก่อนสร้าง scSession');
// ล้มเหลวต้องไม่บล็อกการนับ — ถือเป็นรอบแรกไปก่อน ดีกว่าเปิด session ไม่ได้เลย
assert(/getStockCountGroupState[\s\S]{0,400}\.catch\(function\s*\(err\)\s*\{[\s\S]{0,160}return null;/.test(startBlock),
  'ถามรอบนับไม่สำเร็จต้อง fallback เป็นรอบแรก ไม่ใช่เปิด session ไม่ได้');

// คีย์ทั้งสองฝั่งต้องให้ผลตรงกันเป๊ะ ไม่งั้นรอบ 2 จับคู่ยอดรอบแรกไม่เจอ แล้วกลายเป็นนับปิดยอดซ้ำ
const feKeySrc = htmlLf.match(/^ {4}function scItemKey\([\s\S]*?\n {4}}$/m);
const beKeySrc = backendLf.match(/^function stockCountItemKey\([\s\S]*?\n}$/m);
assert(feKeySrc && beKeySrc, 'ต้องดึงตัวฟังก์ชันคีย์ทั้งสองฝั่งออกมารันได้');
const scItemKey = new Function(feKeySrc[0] + '\nreturn scItemKey;')();
const stockCountItemKey = new Function(beKeySrc[0] + '\nreturn stockCountItemKey;')();
[
  ['Air cylinder', 'CDQ2A20', 'Coil Winding'],
  ['  Air cylinder  ', '  cdq2a20  ', '  COIL WINDING  '],
  ['Bearing', '-', 'Arc Chute'],
  ['Bearing', '', 'Arc Chute'],
  ['Motor', null, ''],
].forEach(function (args) {
  assert.strictEqual(scItemKey.apply(null, args), stockCountItemKey.apply(null, args),
    'คีย์สองฝั่งต้องตรงกัน: ' + JSON.stringify(args));
});
// '-' คือ placeholder ของรุ่นว่าง ไม่ใช่รุ่นจริง — ต้องถือเป็นตัวเดียวกับรุ่นว่าง
assert.strictEqual(stockCountItemKey('Bearing', '-', 'X'), stockCountItemKey('Bearing', '', 'X'));
// ชื่อซ้ำข้ามชีทต้องไม่ถูกจับเป็นชิ้นเดียวกัน (ห้ามจับคู่ข้ามไลน์ ตามกติกาเดิมของระบบ)
assert.notStrictEqual(stockCountItemKey('Bearing', 'B1', 'Coil Winding'), stockCountItemKey('Bearing', 'B1', 'Arc Chute'));

// ── Backend: 1 คน = 1 รอบ ส่งซ้ำต้องทับของตัวเอง ────────────────────────────────
const saveBlock = slice(backendLf, 'function saveStockCountResult(payload)',
  'function getStockCountHistory(payload)', 'saveStockCountResult');
assert(/mine = r;/.test(saveBlock) && /if \(mine\) \{/.test(saveBlock),
  'คนเดิมส่งซ้ำต้องทับแถวของตัวเอง ไม่ใช่ถูกนับเป็นคนที่ 2');
assert(/if \(!isStockCountPending\(r\.values\[idx\.status\]\)\) return;/.test(saveBlock),
  'ทับได้เฉพาะแถวที่ยังไม่ถูกตัดสิน — ที่อนุมัติไปแล้วคือหลักฐาน ห้ามแก้ย้อนหลัง');
assert(/round_no: roundNo/.test(saveBlock) && /var roundNo = groupRows\.length \+ 1;/.test(saveBlock),
  'คนใหม่ต้องได้เลขรอบถัดไป');

const stateBlock = slice(backendLf, 'function getStockCountGroupState(payload)',
  'function getStockCountComparison(payload)', 'getStockCountGroupState');
assert(/var nextRound = myRound > 0 \? myRound : rows\.length \+ 1;/.test(stateBlock),
  'คนที่เคยนับแล้วต้องกลับเข้ารอบเดิมของตัวเอง');
// ต้องทานของคนอื่น ไม่ใช่ทานยอดของตัวเองที่เพิ่งส่งไป
assert(/!== me\) \{ refRow = rows\[j\]; break; \}/.test(stateBlock),
  'ยอดอ้างอิงต้องมาจากรอบล่าสุดที่ไม่ใช่ของตัวเอง');
// แถวที่ยกมาจาก localStorage เป็นหลักฐานเก่า ห้ามถูกนับเป็นรอบทาน
const groupRowsBlock = slice(backendLf, 'function readStockCountGroupRows(data, idx, groupId)',
  'function parseStockCountItems(raw)', 'readStockCountGroupRows');
assert(/indexOf\('archived_'\) === 0\) continue;/.test(groupRowsBlock),
  'แถว archived_* ห้ามถูกนับเป็นรอบนับ');

// ── Backend: ผลเทียบ 3 ถัง ─────────────────────────────────────────────────────
const cmpBlock = slice(backendLf, 'function getStockCountComparison(payload)',
  'function approveStockCountGroup(payload)', 'getStockCountComparison');
assert(/it\.agree = values\.length > 0 && values\.every\(/.test(cmpBlock), 'ต้องเช็คว่าทุกคนนับได้เท่ากันไหม');
assert(/bucket = values\.length === 0 \? 'uncounted'/.test(cmpBlock), 'ต้องแยกถังรายการที่ยังไม่มีใครนับ');
assert(/: !it\.agree \? 'conflict'/.test(cmpBlock), 'นับไม่ตรงกันเอง = conflict ต้องให้ Admin ตัดสิน');
assert(/: it\.matches_system \? 'ok'/.test(cmpBlock) && /: 'adjust';/.test(cmpBlock),
  'ตรงกันแต่ต่างจากระบบ = adjust (ปรับได้)');

// ── Backend: อนุมัติทั้งรอบพร้อมกัน ──────────────────────────────────────────────
// ถ้าอนุมัติทีละคน คนแรกปรับยอดไปแล้ว พอคนที่สองมาอนุมัติจะปรับซ้ำอีกรอบ ยอดเพี้ยน 2 เท่า
const grpBlock = slice(backendLf, 'function approveStockCountGroup(payload)',
  'function approveStockCount(payload)', 'approveStockCountGroup');
assert(/rows\.forEach\(function\s*\(r\)\s*\{[\s\S]{0,300}idx\.status \+ 1\)\.setValue\(decision\)/.test(grpBlock),
  'ต้องประทับสถานะทุกแถวของรอบพร้อมกัน');
assert(grpBlock.indexOf('adjustStockFromCount(') < grpBlock.indexOf('setValue(decision)'),
  'ต้องปรับ Stock ให้เสร็จก่อนประทับสถานะ — ล้มกลางทางแถวยัง pending ให้กดซ้ำได้');
// รายการที่ไม่มีใครนับต้องไม่ถูกแตะ และรายการที่ยอดตรงระบบอยู่แล้วไม่ต้องปรับ
assert(/Number\(r\.counted\) !== Number\(r\.system_qty\)/.test(grpBlock),
  'ปรับเฉพาะรายการที่ยอดสุดท้ายต่างจากระบบ');
assert(/r\.counted !== null && r\.counted !== undefined && r\.counted !== ''/.test(grpBlock),
  'รายการที่ยังไม่ได้ตัดสินยอดต้องไม่ถูกปรับ');

// ── Dispatch ครบทั้ง doGet และ doPost ──────────────────────────────────────────
['getStockCountGroupState', 'getStockCountComparison', 'approveStockCountGroup'].forEach(function (action) {
  assert((backend.match(new RegExp("action === '" + action + "'", 'g')) || []).length === 2,
    'ต้อง dispatch ' + action + ' ทั้ง doGet และ doPost');
});
// คอลัมน์ใหม่ต้องถูกเติมต่อท้ายให้ชีทเก่าเอง ไม่งั้นรอบนับผูกกันไม่ได้
assert(backend.includes("'count_group_id','round_no'"), 'STOCK_COUNT_HEADERS ต้องมีคอลัมน์รอบนับ');

// ── หน้าเว็บ: Admin เห็นผลเทียบก่อนอนุมัติ ─────────────────────────────────────
assert(html.includes('function scOpenComparison(groupId, line)'), 'ต้องมีหน้าผลเทียบสำหรับ Admin');
assert(html.includes("action: 'approveStockCountGroup'"), 'อนุมัติต้องยิงไปที่ระดับรอบ ไม่ใช่ทีละคน');
const cmpSubmit = slice(htmlLf, 'function scSubmitComparison(decision)',
  '// แสดงชื่อ+เหตุผลของรายการที่ปรับ Stock ไม่สำเร็จ', 'scSubmitComparison');
assert(/i\.final !== null && i\.final !== undefined/.test(cmpSubmit),
  'ต้องส่งเฉพาะรายการที่ตัดสินยอดแล้ว');
assert(/loadPartsData\(\{ skipCache: true \}\)/.test(cmpSubmit),
  'ปรับ Stock แล้วต้องโหลดข้อมูลอะไหล่ใหม่ ไม่งั้นตารางค้างยอดเก่า');
// รายการที่ขัดกันยังไม่ได้เลือกยอด ต้องกดอนุมัติไม่ได้ — ไม่งั้นของที่ยังไม่ตัดสินจะถูกข้ามเงียบ ๆ
assert(/unresolved \? 'disabled' : ''/.test(htmlLf), 'ยังตัดสินไม่ครบต้องกดอนุมัติไม่ได้');
// คิวรออนุมัติต้องรวมเป็นรอบ ไม่ใช่โชว์ทีละคน
const pendingBlock = slice(htmlLf, 'function scRenderPending()', 'function scOpenComparison', 'scRenderPending');
assert(/var gid = String\(h\.count_group_id \|\| ''\)\.trim\(\);/.test(pendingBlock),
  'คิวต้องจัดกลุ่มตามรอบนับ');
assert(/'__legacy__' \+ h\.session_id/.test(pendingBlock),
  'แถวเก่าที่ยังไม่มี count_group_id ต้องยังอนุมัติได้ทีละแถวตามเดิม');
assert(/ยังไม่มีคนทาน/.test(pendingBlock),
  'รอบที่มีคนนับคนเดียวต้องเตือน Admin ว่ายังไม่มีคนทาน');

console.log('stock-count-mobile-doublecount: OK');
