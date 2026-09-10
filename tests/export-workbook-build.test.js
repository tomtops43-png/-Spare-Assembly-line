// การประกอบไฟล์ Executive Workbook / CSV / JSON
// จุดที่กันไว้:
//  1) ชุดข้อมูลที่ดึงไม่สำเร็จต้องยังมีชีตของตัวเองที่บอกว่า "ดึงไม่ได้" ไม่ใช่หายไปเงียบๆ
//     (ผู้จัดการเปิดไฟล์เห็นชีตครบแต่ตัวเลขไม่ครบ จะเข้าใจผิดว่าเดือนนั้นไม่มีรายการ)
//  2) ชื่อชีต .xlsx จำกัด 31 ตัวและห้ามอักขระบางตัว ถ้าไม่ล้างชื่อจะเปิดไฟล์ไม่ได้เลย
//  3) CSV ต้องมี BOM ไม่งั้น Excel เปิดภาษาไทยเป็นตัวยึกยือ
//  4) ต้องดึงเรียงทีละชุด ไม่ยิงขนาน 20 คำขอ (Apps Script timeout 30 วิ + ลิมิตคำขอพร้อมกัน)
const { html, grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "2026-09-01", to: "2026-09-10", line: "H9", category: "", txnType: "", groupBy: "month" };',
    'var currentUser = { username: "somchai", role: "leader" };',
    grabVar('XP_DICTIONARY'),
    grabFn('xpSafeSheetName'),
    grabFn('xpTableToCsv'),
    grabFn('xpCsvCell'),
    grabFn('xpFileSuffix'),
    grabFn('xpBuildCover'),
    grabFn('xpBuildDictionary'),
    grabFn('xpBuildAttachmentSheet')
  ]),
  '{ xpSafeSheetName: xpSafeSheetName, xpTableToCsv: xpTableToCsv, xpBuildCover: xpBuildCover, xpBuildDictionary: xpBuildDictionary, xpBuildAttachmentSheet: xpBuildAttachmentSheet, xpFileSuffix: xpFileSuffix, setFilters: function(f) { Object.keys(f).forEach(function(k) { xpFilters[k] = f[k]; }); } }'
);

// ── ชื่อชีต Excel ────────────────────────────────────────────────────
let used = {};
assert.strictEqual(mod.xpSafeSheetName('Master อะไหล่', used), 'Master อะไหล่');
// อักขระที่ Excel ห้ามใช้ในชื่อชีต: [ ] * / \ ? :
const cleaned = mod.xpSafeSheetName('Gv.2 [6 plate]/รอบ*นับ?', used);
assert(!/[\[\]\*\/\\\?:]/.test(cleaned), 'ชื่อชีตต้องไม่มีอักขระที่ Excel ห้าม แต่ได้: ' + cleaned);
// ยาวเกินต้องถูกตัด (Excel จำกัด 31 ตัว)
const long = mod.xpSafeSheetName('ชื่อชีตที่ยาวมากเกินสามสิบเอ็ดตัวอักษรแน่นอนเลยจริงๆ', used);
assert(long.length <= 31, 'ชื่อชีตต้องไม่เกิน 31 ตัว แต่ได้ ' + long.length);
// ชื่อซ้ำต้องถูกเปลี่ยน ไม่งั้น ExcelJS จะโยน error แล้วไฟล์ไม่ออกเลย
used = {};
const a = mod.xpSafeSheetName('PR audit', used);
const b = mod.xpSafeSheetName('PR audit', used);
assert.notStrictEqual(a, b, 'ชื่อชีตซ้ำต้องถูกทำให้ไม่ซ้ำ');

// ── CSV ──────────────────────────────────────────────────────────────
const csv = mod.xpTableToCsv({
  headers: ['ชื่ออะไหล่', 'หมายเหตุ', 'จำนวน'],
  rows: [['น็อต M8', 'ใส่ "เครื่อง" LS-10', 12], ['สายพาน, ยาว', 'บรรทัด\nใหม่', 3]]
});
assert.strictEqual(csv.charCodeAt(0), 0xFEFF, 'CSV ต้องขึ้นต้นด้วย BOM ไม่งั้น Excel อ่านภาษาไทยเพี้ยน');
assert(csv.indexOf('"ใส่ ""เครื่อง"" LS-10"') > -1, 'เครื่องหมายคำพูดในค่าต้องถูก escape เป็น ""');
assert(csv.indexOf('"สายพาน, ยาว"') > -1, 'ค่าที่มีคอมมาต้องอยู่ในเครื่องหมายคำพูด');
assert(csv.indexOf('\r\n') > -1, 'ต้องขึ้นบรรทัดใหม่แบบ CRLF ให้ Excel บน Windows อ่านได้');

// ── หน้าปกรายงาน ─────────────────────────────────────────────────────
const cover = mod.xpBuildCover([
  { ds: { name: 'Master อะไหล่', sheetName: 'Master อะไหล่' }, table: { headers: ['a'], rows: [[1], [2], [3]] } },
  { ds: { name: 'PR audit', sheetName: 'PR audit' }, error: 'ไม่มีสิทธิ์' }
]);
const coverText = cover.rows.map(function(r) { return r.join(' | '); }).join('\n');
assert(/ช่วงข้อมูล \| 2026-09-01  ถึง  2026-09-10/.test(coverText), 'ปกต้องระบุช่วงข้อมูลที่กรองไว้');
assert(/ไลน์ \| H9/.test(coverText), 'ปกต้องระบุไลน์ที่กรองไว้');
assert(/ผู้จัดทำ \| somchai/.test(coverText), 'ปกต้องระบุผู้จัดทำ');
assert(/Master อะไหล่ \| 3/.test(coverText), 'ปกต้องบอกจำนวนแถวของแต่ละชีต');
assert(/PR audit \| ดึงข้อมูลไม่ได้: ไม่มีสิทธิ์/.test(coverText), 'ชุดที่ดึงไม่ได้ต้องระบุไว้บนปก ไม่ใช่หายไปเงียบๆ');
assert(/ความครบถ้วนของข้อมูล \| มี 1 ชุดที่ดึงไม่สำเร็จ/.test(coverText), 'ปกต้องสรุปความครบถ้วนให้คนอ่านเห็นทันที');
assert(/ไม่มีตัวเลขที่สร้างขึ้นเอง/.test(coverText), 'ปกต้องระบุที่มาข้อมูลว่าดึงสดจากระบบ');

// โหมดลูกค้าต้องเขียนกำกับไว้บนปก
mod.setFilters({ customerMode: true });
const coverCust = mod.xpBuildCover([]).rows.map(function(r) { return r.join(' | '); }).join('\n');
assert(/โหมดลูกค้า \| เปิด/.test(coverCust), 'ไฟล์โหมดลูกค้าต้องมีคำกำกับบนปก');
mod.setFilters({ customerMode: false });

// ── พจนานุกรมข้อมูล ──────────────────────────────────────────────────
const dict = mod.xpBuildDictionary([
  { name: 'Master อะไหล่', table: { headers: ['ชื่อรายการ', 'คงเหลือ', 'Min', 'คอลัมน์แปลกใหม่'] } },
  { name: 'Log', table: { headers: ['ยอดหลัง'] } }
]);
assert.deepStrictEqual(dict.headers, ['ชีตในไฟล์นี้', 'ชื่อคอลัมน์', 'คำอธิบาย / วิธีคำนวณ']);
const dictMap = {};
dict.rows.forEach(function(r) { dictMap[r[1]] = r[2]; });
assert(/จุดสั่งซื้อ/.test(dictMap['Min']), 'คอลัมน์สำคัญต้องมีคำอธิบายจริง ไม่ใช่ข้อความ fallback');
assert(/น้อยกว่าเท่านั้น/.test(dictMap['Min']), 'ต้องอธิบายว่า Min เทียบแบบน้อยกว่า (ตรงกับกฎธุรกิจในโค้ด)');
assert(/ตรวจย้อนหลัง/.test(dictMap['ยอดหลัง']), 'คอลัมน์ยอดหลังต้องอธิบายว่าใช้ทำอะไร');
assert.strictEqual(dictMap['คอลัมน์แปลกใหม่'], 'คอลัมน์ตรงจากชีตต้นทางของระบบ', 'คอลัมน์ที่ยังไม่มีคำอธิบายต้องมีข้อความ fallback ไม่ใช่ช่องว่าง');
// ทุกคอลัมน์ในทุกชีตต้องมีบรรทัดของตัวเอง
assert.strictEqual(dict.rows.length, 5);

// ── ชีตลิงก์ไฟล์แนบ ──────────────────────────────────────────────────
const attach = mod.xpBuildAttachmentSheet({
  master: [{ __sheet: 'H9', no: '1', name: 'สายพาน', model: 'B-77', image_main_url: 'https://drive/img1', drawing_url: 'https://drive/dwg1', quotation_url: 'https://drive/quo1' }],
  orderRequests: [{ request_id: 'RQ1', item_name: 'น็อต', model: 'M8', attachment_url: 'https://drive/rq1' }],
  miscExpenses: [{ expense_id: 'EX1', item_name: 'ประแจ', receipt_url: 'https://drive/rc1' }],
  purchaseHistory: [{ 'History ID': 'PH1', 'Part Name': 'เบรกเกอร์', 'Model / Part No.': 'BK-10', 'Attachment URL': 'https://drive/ph1' }]
}, []);
const links = attach.rows.map(function(r) { return r[6]; });
assert.strictEqual(links.length, 6, 'ต้องเก็บทุกลิงก์ที่มีค่า (รูป/Drawing/ใบเสนอราคา/แนบคำขอ/บิล/แนบซื้อ)');
assert(links.indexOf('https://drive/img1') > -1);
assert(links.indexOf('https://drive/quo1') > -1);
// รายการที่ไม่มีลิงก์ต้องไม่สร้างแถวเปล่า
assert(attach.rows.every(function(r) { return r[6]; }), 'ห้ามมีแถวที่ลิงก์ว่าง');

// ── ชื่อไฟล์ต้องบอกได้ว่าเป็นข้อมูลชุดไหน ─────────────────────────────
mod.setFilters({ line: 'Lug&Screw', from: '2026-09-01', to: '2026-09-10', customerMode: true });
const suffix = mod.xpFileSuffix();
assert(/Lug-Screw/.test(suffix), 'ชื่อไฟล์ต้องมีไลน์');
assert(/2026-09-01_ถึง_2026-09-10/.test(suffix), 'ชื่อไฟล์ต้องมีช่วงวันที่');
assert(/customer/.test(suffix), 'ไฟล์โหมดลูกค้าต้องดูออกจากชื่อไฟล์ กันส่งไฟล์ผิดชุดให้ลูกค้า');
assert(!/[\\\/:*?"<>|]/.test(suffix), 'ชื่อไฟล์ต้องไม่มีอักขระที่ระบบไฟล์ห้าม');

// ── โครงชีตของ Executive Workbook ────────────────────────────────────
const wbSrc = grabFn('xpRunWorkbook');
['00 ปกรายงาน', '01 สรุปผู้บริหาร', '02 เทรนด์ตามช่วงเวลา', '03 สรุปตามไลน์', '98 ลิงก์ไฟล์แนบ', '99 พจนานุกรมข้อมูล'].forEach(function(name) {
  assert(wbSrc.indexOf(name) > -1, 'Executive Workbook ต้องมีชีต ' + name);
});
assert(/ดึงข้อมูลชุดนี้ไม่สำเร็จ/.test(wbSrc), 'ชุดที่ดึงไม่ได้ต้องยังมีชีตของตัวเองที่บอกสาเหตุ');

// ── ต้องดึงเรียงทีละชุด ไม่ยิงขนาน ────────────────────────────────────
const seqSrc = grabFn('xpFetchDatasetsSequential');
assert(!/Promise\.all/.test(seqSrc), 'ห้ามใช้ Promise.all ดึง 20 ชุดพร้อมกัน — Apps Script จะทิ้งคำขอบางส่วน');
assert(/\.then\(next\)/.test(seqSrc), 'ต้องดึงต่อกันเป็นลูกโซ่');
assert(/results\.push\(\{ ds: ds, table: null, error/.test(seqSrc), 'ชุดที่พังต้องถูกบันทึกแล้วไปต่อ ไม่ล้มทั้งงาน');
const ctxSrc = grabFn('xpContext');
assert(!/Promise\.all/.test(ctxSrc), 'xpContext ก็ต้องดึงเรียงทีละชุด');
assert(/ctx\.failed\[key\]/.test(ctxSrc), 'ชุดที่ดึงไม่สำเร็จต้องถูกจดไว้ใน ctx.failed เพื่อให้ KPI แสดงว่า "ดึงไม่ได้"');

// ── ExcelJS ต้องโหลดตอนใช้ และต้องมีทางถอย ───────────────────────────
const excelSrc = grabFn('xpLoadExcelJs');
assert(/cdn\.jsdelivr\.net\/npm\/exceljs@4\.4\.0/.test(excelSrc), 'ต้องปักหมุดเวอร์ชัน ExcelJS ไม่ใช่ latest');
assert(!/<script src=[^>]*exceljs/.test(html), 'ห้ามโหลด ExcelJS ตอนเปิดเว็บ — คนส่วนใหญ่ไม่ได้เข้าหน้านี้ทุกวัน');
const writeSrc = grabFn('xpWriteXlsxStyled');
assert(/window\.XLSX/.test(writeSrc), 'ถ้า ExcelJS โหลดไม่ได้ ต้องถอยไปใช้ SheetJS ที่มีอยู่แล้ว');
assert(/styled: false/.test(writeSrc), 'ต้องบอกผู้ใช้ได้ว่าไฟล์ที่ได้ไม่มีการจัดรูปแบบ');
assert(/rowCount <= 5000/.test(writeSrc), 'ชีตใหญ่ต้องไม่ลงลายเส้นทีละแถว (ช้าและไฟล์บวม)');

console.log('export-workbook-build: OK');
