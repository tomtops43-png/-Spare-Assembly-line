const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── ค่าใช้จ่ายสิ้นเปลืองต้องไหลเข้ากราฟ % ต้นทุน ────────────────────────────────
// ของจิปาถะที่ซื้อนอกระบบ (น็อต/เครื่องมือ) ไม่มีใน master list จึงไม่มีวันโผล่ในยอดเบิก
// ที่ Dashboard ใช้คิดรายจ่าย — ต้องมีท่อของตัวเองที่แปลงเป็นแถวหน้าตาเดียวกัน
const src = [
  grab(/^ {4}function fccN\(v\) \{.*\}$/m, 'fccN'),
  grab(/^ {4}function fccMonthKey\(raw\) \{[\s\S]*?\n {4}\}/m, 'fccMonthKey'),
  grab(/^ {4}function fccLineOf\(raw\) \{[\s\S]*?\n {4}\}/m, 'fccLineOf'),
  grab(/^ {4}function fccMiscExpenseRows\(list\) \{[\s\S]*?\n {4}\}/m, 'fccMiscExpenseRows')
].join('\n');
const fccMiscExpenseRows = new Function(src + '\nreturn fccMiscExpenseRows;')();

const rows = fccMiscExpenseRows([
  { month: '2026-09', date: '2026-09-03', line: 'H9', category: 'น็อต/สกรู/ฮาร์ดแวร์', item_name: 'น็อต M8', total_amount: 480, vendor: 'ร้านไทยวัสดุ' },
  { month: '', date: '2026-09-05', line: 'ส่วนกลาง', category: 'เครื่องมือช่าง', item_name: 'ประแจ', total_amount: 1250 },
  { month: new Date('2026-08-31T17:00:00.000Z'), date: '', line: 'Lug&Screw', category: 'วัสดุสิ้นเปลือง', item_name: 'เทปพันสายไฟ', total_amount: 90 },
  { month: '2026-09', date: '2026-09-07', line: 'H9', category: 'อื่นๆ', item_name: 'ของแถม', total_amount: 0 }
]);
assert.strictEqual(rows.length, 3, 'ยอด 0 บาทต้องไม่ถูกนับเป็นรายจ่าย');
assert.strictEqual(rows[0].amount, 480, 'amount ต้องมาจาก total_amount ตรงๆ');
assert.strictEqual(rows[0].month, '2026-09');
assert.strictEqual(rows[1].month, '2026-09', 'ไม่มี month ต้อง fallback ไปอ่านจาก date');
// Google Sheets ชอบแปลงคอลัมน์ Month เป็น Date object — ต้องอ่านที่โซนเวลาไทย ไม่ใช่ UTC
// (31 ส.ค. 17:00Z = 1 ก.ย. เวลาไทย ต้องนับเป็นรอบเดือนกันยายน)
assert.strictEqual(rows[2].month, '2026-09', 'Date object ต้องถูกอ่านที่ Asia/Bangkok');
// 'ส่วนกลาง' ต้องไม่ถูกจับไปรวมกับไลน์ไหน — ตั้งใจให้เห็นเฉพาะมุมมอง "ทุกไลน์ (รวม)"
assert.strictEqual(rows[1].line, 'ส่วนกลาง');
assert.strictEqual(rows[0].line, 'H9');
assert.strictEqual(rows[2].line, 'Lug&Screw');

// ── ต้องเสียบเข้าท่อรายจ่ายทั้งสองทาง: กราฟ % ต้นทุน และระบบคุมงบ PR ──────────
assert(/miscRows\.filter\(function\(r\) \{ return isAll \? true : r\.line === line; \}\)[\s\S]{0,320}spendByMonth\[r\.month\] = \(spendByMonth\[r\.month\] \|\| 0\) \+ r\.amount;/.test(htmlLf),
  'กราฟ % ต้นทุนต้องบวกค่าใช้จ่ายสิ้นเปลืองเข้า spendByMonth');
assert(/fccMiscExpenseRows\(A\.misc\)\.forEach\(function\(r\) \{[\s\S]{0,200}spentBML\[r\.month\]\[r\.line\]/.test(htmlLf),
  'ระบบคุมงบ PR ต้องนับค่าใช้จ่ายสิ้นเปลืองเข้า used ของเดือนนั้นด้วย');
assert(htmlLf.includes("var pMisc = requestApi(withAuthPayload({ action: 'getMiscExpenses' }))"),
  'Dashboard ต้องโหลด getMiscExpenses มาพร้อมชุดข้อมูลอื่น');
assert(htmlLf.includes('Promise.all([pReq, pLog, pHist, pProdVolume, pProdConfig, pMisc])'));
assert(htmlLf.includes('fccAsyncCache.misc = res[5];') && htmlLf.includes('fccLast.Misc = A.misc;'));
// แท่งกราฟต้องแยกสีให้เห็นว่าเงินหมดไปกับของในระบบหรือของซื้อนอกระบบ (stack เดียวกัน = เทียบกับงบได้)
assert(htmlLf.includes("label: 'ของสิ้นเปลือง (ซื้อนอกระบบ)'") && htmlLf.includes("stack: 'spend'") && htmlLf.includes("stack: 'budget'"));
assert(htmlLf.includes('return d.expenseParts;'), 'แท่งอะไหล่ต้องใช้ยอดเฉพาะอะไหล่ ไม่ใช่ยอดรวม (ไม่งั้นซ้อนแล้วเกินจริง)');

// ── หน้าเว็บ + ฟอร์ม ─────────────────────────────────────────────────────────
assert(htmlLf.includes('id="miscExpensePage"') && htmlLf.includes('id="tabMiscExpense"'));
assert(/<button id="tabMiscExpense"[\s\S]{0,600}?>ค่าใช้จ่ายสิ้นเปลือง<\/span>/.test(htmlLf), 'ต้องมีปุ่มเมนูในกลุ่มจัดซื้อ');
assert(htmlLf.includes("'misc-expense': 'procure'"), 'ต้องไฮไลต์กลุ่ม "จัดซื้อ" ตอนอยู่หน้านี้');
// 1 บิลกรอกได้หลายรายการหลายหมวด — ยอดรวมทั้งบิลคิดเองจากทุกแถวและห้ามแก้
assert(htmlLf.includes('id="miscExpenseRows"') && htmlLf.includes('id="miscExpenseAddRow"'), 'ต้องมีที่เพิ่มรายการหลายแถว');
assert(/<output id="miscExpenseGrandTotal"/.test(htmlLf), 'ยอดรวมทั้งบิลต้องเป็น output ที่แก้เองไม่ได้ ไม่ใช่ input');
assert(htmlLf.includes('id="miscExpenseReceiptFile"') && htmlLf.includes('capture="environment"'), 'ต้องถ่ายรูปบิลจากมือถือได้');
assert(htmlLf.includes('<option value="ส่วนกลาง">'), 'ต้องเลือก "ส่วนกลาง" ได้สำหรับของที่ใช้ร่วมหลายไลน์');
// ต้องซ่อนหน้านี้ทุกครั้งที่สลับไปหน้าอื่น ไม่งั้นหน้าซ้อนกัน
assert((htmlLf.match(/^ {6}mxHidePage\(\);$/gm) || []).length >= 10, 'ทุก switch page ต้องเรียก mxHidePage()');

// ── สิทธิ์: Leader บันทึกได้ / Admin เท่านั้นที่แก้-ลบ ────────────────────────
assert(/function canAddMiscExpense\(\)[\s\S]{0,220}role === 'admin' \|\| role === 'leader'/.test(htmlLf));
assert(/function canManageMiscExpense\(\)[\s\S]{0,200}=== 'admin';/.test(htmlLf));
assert(/function requireMiscExpenseEditor\(payload\)[\s\S]{0,420}role !== 'admin' && role !== 'leader'/.test(backend));
assert(/function requireMiscExpenseAdmin\(payload\)[\s\S]{0,320}!== 'admin'/.test(backend));
assert(backend.includes("var user = requireMiscExpenseAdmin({ authToken: payload.authToken });"), 'แก้/ลบต้องผ่าน Admin gate');

// ── Backend: ชีต + routing ───────────────────────────────────────────────────
assert(backend.includes("SPARE_APP_CONFIG.miscExpenseSheetName = SPARE_APP_CONFIG.miscExpenseSheetName || 'MiscExpenses'"));
assert(backend.includes("var MISC_EXPENSE_HEADERS = ['Expense ID', 'Date', 'Month', 'Line', 'Category', 'Item Name', 'Qty', 'Unit', 'Unit Price', 'Total Amount', 'Vendor', 'Receipt No', 'Receipt URL', 'Paid By', 'Remark', 'Deleted', 'Created By', 'Created At', 'Updated By', 'Updated At', 'Bill ID'];"));
['getMiscExpenses', 'addMiscExpense', 'updateMiscExpense', 'deleteMiscExpense'].forEach(function(action) {
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(e.parameter), e);"), action + ' ต้องเรียกผ่าน GET ได้');
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(body), e);"), action + ' ต้องเรียกผ่าน POST ได้');
});
assert(backend.includes("if (action === 'uploadMiscExpenseReceipt') return respond(uploadMiscExpenseReceipt(body), e);"));
// เงินที่จ่ายไปแล้วต้องตามรอยย้อนหลังได้เสมอ — ลบแบบ soft delete + มี audit log
assert(/function deleteMiscExpense\(payload\)[\s\S]{0,900}sheet\.getRange\(rowIndex \+ 1, 16\)\.setValue\(true\);/.test(backend), 'ต้องลบแบบ soft delete ไม่ใช่ลบแถวทิ้ง');
assert(/appendMiscExpenseAudit\(user\.username, expenseId, 'DELETE'/.test(backend));
assert(backend.includes("if (!reason) throw new Error('กรุณาระบุเหตุผลการลบ');"));
// กันกรอกวันที่อนาคต (พิมพ์ปีผิดแล้วยอดไปโผล่เดือนที่ยังมาไม่ถึง หาไม่เจอ)
assert(backend.includes("throw new Error('บันทึกวันที่ล่วงหน้าไม่ได้');"));

console.log('Misc expense ledger checks passed');

// ── 1 บิล = หลายรายการหลายหมวด (กรอกรวดเดียว ไม่ต้องเข้าออกทีละชิ้น) ──────────
const rowSrc = [
  grab(/^ {4}function mxRowAmount\(row\) \{[\s\S]*?\n {4}\}/m, 'mxRowAmount'),
  grab(/^ {4}function mxIsBlankRow\(row\) \{[\s\S]*?\n {4}\}/m, 'mxIsBlankRow')
].join('\n');
const mxRowAmount = new Function(rowSrc + '\nreturn mxRowAmount;')();
const mxIsBlankRow = new Function(rowSrc + '\nreturn mxIsBlankRow;')();

// เป็นเงินของแต่ละแถว = จำนวน × ราคา/หน่วย เหมือนบิลร้านค้า (5 เมตร × 14 = 70)
assert.strictEqual(mxRowAmount({ qty: 5, unit_price: 14 }), 70);
assert.strictEqual(mxRowAmount({ qty: 3, unit_price: 9 }), 27);
assert.strictEqual(mxRowAmount({ qty: '1', unit_price: '16' }), 16, 'ค่าจาก input เป็น string ต้องคิดได้');
assert.strictEqual(mxRowAmount({ qty: 3, unit_price: 0.335 }), 1.01, 'ต้องปัดเป็น 2 ตำแหน่ง ไม่ปล่อยทศนิยมลอย');
assert.strictEqual(mxRowAmount({ qty: 1, unit_price: '' }), 0, 'ยังไม่กรอกราคา = 0 ไม่ใช่ NaN');
assert.strictEqual(mxRowAmount({ qty: 0, unit_price: 50 }), 0, 'จำนวน 0 ต้องไม่กลายเป็นยอดติดลบ/NaN');
// รวมทั้งบิลตามตัวอย่างจริง: 70 + 75 + 27 + 16 + 16 + 16 = 220
const billRows = [
  { qty: 5, unit_price: 14 }, { qty: 5, unit_price: 15 }, { qty: 3, unit_price: 9 },
  { qty: 1, unit_price: 16 }, { qty: 1, unit_price: 16 }, { qty: 1, unit_price: 16 }
];
assert.strictEqual(billRows.reduce((s, r) => s + mxRowAmount(r), 0), 220);
// แถวที่เผลอกดเพิ่มแล้วไม่ได้กรอกอะไรเลย ต้องถูกทิ้ง ไม่ใช่เด้ง error ให้ลบเอง
assert.strictEqual(mxIsBlankRow({ item_name: '', unit_price: '', category: '' }), true);
assert.strictEqual(mxIsBlankRow({ item_name: 'น็อต M8', unit_price: '', category: '' }), false);

// ยอดรวมทั้งบิลต้องคำนวณจากทุกแถว ไม่ใช่ให้พิมพ์เอง
assert(/function mxRecalcTotals\(\)[\s\S]{0,900}grand \+= amount;/.test(htmlLf));
assert(/mxCollectItems\(\)[\s\S]{0,400}total_amount: mxRowAmount\(row\)/.test(htmlLf), 'ยอดที่ส่งขึ้นเซิร์ฟเวอร์ต้องมาจาก mxRowAmount ของแถวนั้น');
assert(htmlLf.includes("action: 'addMiscExpenseBatch', items: items"), 'บันทึกทั้งบิลต้องส่งเป็นชุดครั้งเดียว');
// วาดแถวใหม่ทุกคีย์ = เคอร์เซอร์เด้ง ต้องอัปเดตแค่ยอดระหว่างพิมพ์
assert(/function mxOnRowInput\(e\)[\s\S]{0,700}mxRecalcTotals\(\);/.test(htmlLf));
assert(!/function mxOnRowInput\(e\)[\s\S]{0,700}mxRenderRows\(\);/.test(htmlLf), 'ห้าม re-render ระหว่างพิมพ์');
// แก้ไขทีละรายการ (แถวเดียว) — ต้องซ่อนปุ่มเพิ่มแถวไม่ให้แตกบิลตอนแก้
assert(/var addBtn = mxEl\('miscExpenseAddRow'\);[\s\S]{0,140}classList\.toggle\('hidden', single\)/.test(htmlLf));

// ── Backend: บันทึกเป็นชุด 1 บิล ────────────────────────────────────────────
assert(backend.includes("if (action === 'addMiscExpenseBatch') return respond(addMiscExpenseBatch(e.parameter), e);"));
assert(backend.includes("if (action === 'addMiscExpenseBatch') return respond(addMiscExpenseBatch(body), e);"));
assert(/function addMiscExpenseBatch\(payload\)[\s\S]{0,600}requireMiscExpenseEditor/.test(backend), 'บันทึกทั้งบิลต้องผ่าน gate เดียวกับบันทึกเดี่ยว');
// ตรวจครบทุกแถวก่อนค่อยเขียน — กันบิลเข้าไปครึ่งใบแล้วเจอแถวเสียตรงกลาง
assert(/var fields = rawItems\.map\(function\(item, index\) \{[\s\S]{0,900}throw new Error\('รายการที่ ' \+ \(index \+ 1\)/.test(backend));
assert(/var rows = fields\.map[\s\S]{0,700}sheet\.getRange\(sheet\.getLastRow\(\) \+ 1, 1, rows\.length, MISC_EXPENSE_HEADERS\.length\)\.setValues\(rows\);/.test(backend),
  'ต้องเขียนทีเดียวทั้งบิล ไม่ใช่ appendRow ทีละแถว');
assert(backend.includes("throw new Error('ต้องมีอย่างน้อย 1 รายการ');") && backend.includes('บันทึกได้สูงสุด 50 รายการต่อ 1 บิล'));
// ทุกแถวในบิลเดียวกันต้องได้ Bill ID เดียวกัน และแก้ไขภายหลังต้องไม่ทำ Bill ID หาย
assert(/var billId = 'MEB-' \+ Utilities\.getUuid\(\);/.test(backend));
assert(/appendMiscExpenseAudit\(user\.username, billId, 'CREATE_BILL'/.test(backend));
assert(backend.includes("user.username || '', timestamp, existing[20] || ''"), 'แก้ไขแถวเดิมต้องคง Bill ID ไว้');
// ชีตที่สร้างไปก่อนมีคอลัมน์ Bill ID ต้องเติมหัวคอลัมน์ให้เอง ไม่ใช่ migrate ทั้งชีต
assert(/if \(sheet\.getLastColumn\(\) < MISC_EXPENSE_HEADERS\.length\) \{/.test(backend));

console.log('Misc expense multi-item bill checks passed');

// ── การ์ดบน Dashboard (มุมมองผู้จัดการแผนก) ──────────────────────────────────
// ค่าใช้จ่ายพวกนี้ถูกบวกในกราฟ Cost Ratio อยู่แล้ว แต่ผู้จัดการต้องเห็นด้วยว่า
// "เงินหมดไปกับหมวดไหน ไลน์ไหน ร้านไหน" ไม่ใช่เห็นแค่ยอดรวมก้อนเดียว
assert(htmlLf.includes('id="fccMiscCard"') && htmlLf.includes('id="fccMiscSummary"') && htmlLf.includes('id="fccMiscTable"'),
  'ต้องมีการ์ดค่าใช้จ่ายสิ้นเปลืองบน Dashboard');
assert(htmlLf.includes('id="fccMiscTrendCanvas"') && htmlLf.includes('id="fccMiscBreakdown"'));
assert(htmlLf.includes('id="fccMiscLineFilter"') && htmlLf.includes('id="fccMiscRangeFilter"'), 'ต้องกรองตามไลน์และช่วงเวลาได้');
assert(/fccRenderRepairCost\(\);\r?\n\s+fccRenderMiscExpense\(\);/.test(htmlLf), 'ต้องถูกเรียกตอน render Dashboard');
assert(/\['fccMiscLineFilter', 'fccMiscRangeFilter'\]\.forEach\(function\(id\) \{[\s\S]{0,200}addEventListener\('change', fccRenderMiscExpense\)/.test(htmlLf),
  'เปลี่ยนตัวกรองต้องคำนวณใหม่จากข้อมูลที่โหลดไว้ ไม่ยิง API ซ้ำ');
// ข้อมูลยังโหลดไม่เสร็จต้องคง skeleton ไว้ ไม่ใช่วาดการ์ดเปล่า
assert(/function fccRenderMiscExpense\(\)[\s\S]{0,900}if \(!Array\.isArray\(fccLast\.Misc\)\) return;/.test(htmlLf));
// Chart.js instance ต้อง destroy ก่อนวาดใหม่ ไม่งั้น canvas ซ้อนกันจน hover เพี้ยน
assert(/if \(fccMiscTrendChart\) \{ fccMiscTrendChart\.destroy\(\); fccMiscTrendChart = null; \}/.test(htmlLf));
// KPI ที่ผู้จัดการใช้ตัดสินใจ: สัดส่วนของซื้อนอกระบบเทียบรายจ่ายอะไหล่จริงในขอบเขตเดียวกัน
assert(/partsSpend \+= r\.amount;[\s\S]{0,300}var sharePct = totalSpend > 0 \? \(grand \/ totalSpend \* 100\) : null;/.test(htmlLf));
assert(htmlLf.includes("miscCard('สัดส่วนต่อรายจ่ายอะไหล่ทั้งหมด'"));
// แถวข้อมูลต้องมีสถานะรูปบิล เพื่อเตือนว่ายอดไหนตรวจย้อนหลังไม่ได้
assert(/hasReceipt: !!String\(r\.receipt_url \|\| ''\)\.trim\(\),/.test(htmlLf));
assert(htmlLf.includes('⚠️ ยังไม่แนบรูปบิล '), 'ต้องเตือนรายการที่ไม่มีบิลแนบ');
// "ส่วนกลาง" ต้องอธิบายในการ์ดว่าทำไมไม่ถูกปันเข้าไลน์ ไม่ใช่ปล่อยให้เดาเอง
assert(htmlLf.includes('ไม่ถูกปันเข้าไลน์ใดไลน์หนึ่งและไม่ตัดงบ PR ของไลน์'));
// Export CSV ต้องมี BOM ไม่งั้น Excel อ่านภาษาไทยเป็นตัวยึกยือ
assert(/function fccExportMiscExpenseCsv\(\)/.test(htmlLf) && htmlLf.indexOf("misc-expense-' + new Date().toISOString()") > -1,
  'ต้อง Export CSV ของค่าใช้จ่ายสิ้นเปลืองได้');
assert(htmlLf.indexOf("new Blob(['" + String.fromCharCode(92) + "ufeff' + lines.join") > -1, 'CSV ต้องมี BOM ไม่งั้น Excel อ่านภาษาไทยเป็นตัวยึกยือ');
assert(htmlLf.includes("items.push(['scroll:fccMiscCard'"), 'ต้องมีปุ่มลัดในเมนูลอยให้เลื่อนไปการ์ดนี้');

console.log('Misc expense dashboard card checks passed');
