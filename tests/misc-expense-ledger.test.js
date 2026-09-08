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
// ยอดรวมคือช่องที่คนกรอกจริง (บิลร้านค้ามีแต่ยอดรวม) ราคาต่อหน่วยให้ระบบคิดย้อนเอง
assert(htmlLf.includes('id="miscExpenseTotal"') && htmlLf.includes('ยอดรวม (บาท)'));
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
assert(backend.includes("var MISC_EXPENSE_HEADERS = ['Expense ID', 'Date', 'Month', 'Line', 'Category', 'Item Name', 'Qty', 'Unit', 'Unit Price', 'Total Amount', 'Vendor', 'Receipt No', 'Receipt URL', 'Paid By', 'Remark', 'Deleted', 'Created By', 'Created At', 'Updated By', 'Updated At'];"));
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
