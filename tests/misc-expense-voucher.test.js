const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── ใบเบิกจ่าย / 付款申请单 — ซื้อของด้านนอกแล้วเอาบิลมาเบิกบริษัท ────────────────
const src = [
  grab(/^ {4}function escHtml\(str\) \{[\s\S]*?\n {4}\}/m, 'escHtml'),
  grab(/^ {4}function mxDateLabel\(value\) \{[\s\S]*?\n {4}\}/m, 'mxDateLabel'),
  grab(/^ {4}var MX_VOUCHER_COMPANY_TH[\s\S]*?\n {4}var MX_VOUCHER_LINE_DEPT = [^\n]*/m, 'ค่าคงที่ใบเบิก'),
  grab(/^ {4}function mxBahtText\(amount\) \{[\s\S]*?\n {4}\}/m, 'mxBahtText'),
  grab(/^ {4}function mxVoucherMoney\(n\) \{[\s\S]*?\n {4}\}/m, 'mxVoucherMoney'),
  grab(/^ {4}function mxVoucherData\(rows\) \{[\s\S]*?\n {4}\}/m, 'mxVoucherData'),
  grab(/^ {4}function mxVoucherHtml\(d, appDate\) \{[\s\S]*?\n {4}\}/m, 'mxVoucherHtml')
].join('\n');
const api = new Function(src + '\nreturn { mxBahtText, mxVoucherData, mxVoucherHtml };')();

// ตัวเขียนต้องตรงกับ =BAHTTEXT() ของ Excel ที่ฟอร์มเดิมใช้
assert.strictEqual(api.mxBahtText(56014.5), 'ห้าหมื่นหกพันสิบสี่บาทห้าสิบสตางค์', 'ยอดตัวอย่างจากฟอร์มจริง');
assert.strictEqual(api.mxBahtText(0), 'ศูนย์บาทถ้วน');
assert.strictEqual(api.mxBahtText(1), 'หนึ่งบาทถ้วน');
assert.strictEqual(api.mxBahtText(11), 'สิบเอ็ดบาทถ้วน');
assert.strictEqual(api.mxBahtText(21), 'ยี่สิบเอ็ดบาทถ้วน');
assert.strictEqual(api.mxBahtText(101), 'หนึ่งร้อยเอ็ดบาทถ้วน');
assert.strictEqual(api.mxBahtText(1250), 'หนึ่งพันสองร้อยห้าสิบบาทถ้วน');
assert.strictEqual(api.mxBahtText(0.25), 'ยี่สิบห้าสตางค์');
assert.strictEqual(api.mxBahtText(2500000), 'สองล้านห้าแสนบาทถ้วน');
assert.strictEqual(api.mxBahtText(11000000), 'สิบเอ็ดล้านบาทถ้วน');

// ทั้งบิลออกเป็นใบเดียว: แถว VAT ไปลงช่อง Vat ไม่ใช่รายการสินค้า, ค่าจัดส่งยังเป็นรายการ
const bill = [
  { expense_id: 'E1', bill_id: 'B1', date: '2026-09-03', line: 'H9', category: 'น็อต/สกรู/ฮาร์ดแวร์', item_name: 'น็อต M8', qty: 10, unit: 'ตัว', unit_price: 12, total_amount: 120, vendor: 'ร้านไทยวัสดุ', receipt_no: 'INV-01', created_by: 'สมชาย' },
  { expense_id: 'E2', bill_id: 'B1', date: '2026-09-03', line: 'H9', category: 'ค่าจัดส่ง', item_name: 'ค่าจัดส่ง', qty: 1, unit: 'ครั้ง', total_amount: 50, vendor: 'ร้านไทยวัสดุ' },
  { expense_id: 'E3', bill_id: 'B1', date: '2026-09-03', line: 'H9', category: 'ภาษีมูลค่าเพิ่ม (VAT)', item_name: 'ภาษีมูลค่าเพิ่ม 7%', qty: 1, unit: 'ครั้ง', total_amount: 11.9, vendor: 'ร้านไทยวัสดุ' }
];
const d = api.mxVoucherData(bill);
assert.strictEqual(d.items.length, 2, 'VAT ต้องไม่ถูกนับเป็นรายการสินค้า');
assert.strictEqual(d.subtotal, 170);
assert.strictEqual(d.vat, 11.9);
assert.strictEqual(d.vatLabel, '7%');
assert.strictEqual(d.total, 181.9);
assert.strictEqual(d.totalText, 'หนึ่งร้อยแปดสิบเอ็ดบาทเก้าสิบสตางค์');
assert.strictEqual(d.vendor, 'ร้านไทยวัสดุ');
assert.strictEqual(d.requester, 'สมชาย', 'ผู้ยื่นเบิก = คนที่บันทึกบิล');
assert.strictEqual(d.chargedDept, 'H9 Assembly Production Line', 'ไลน์ H9 ต้องติ๊กช่อง H9 Assembly');
// ไม่มี unit_price (บันทึกจากฟอร์มหลายแถว) ต้องคิดจากยอดรวม/จำนวน
assert.strictEqual(api.mxVoucherData([{ item_name: 'ประแจ', qty: 4, total_amount: 1000 }]).items[0].unitPrice, 250);
// ไลน์ที่ไม่รู้ว่าตรงแผนกไหน ปล่อยให้ติ๊กเอง ดีกว่าติ๊กผิด
assert.strictEqual(api.mxVoucherData([{ line: 'Arc Chute', item_name: 'x', qty: 1, total_amount: 1 }]).chargedDept, '');

const out = api.mxVoucherHtml(d, '12/10/2026');
['ใบเบิกจ่าย/ 付款申请单', '津荣电子科技（泰国）有限公司', 'ชื่อเจ้าหนี้:', '付款方式(จ่าย)', 'ตัวเขียน大写：', 'ยอดรวม合计',
  'Charged Department / แผนกที่เกิดค่าใช้จ่าย / 费用承担部门', 'แผนกบัญชี', 'กรรมการผู้จัดการ', 'ผู้รับเงิน', 'หัวหน้าแผนก 部门核准人：'
].forEach(function(t) { assert(out.includes(t), 'ใบเบิกต้องมีหัวข้อตามฟอร์มกระดาษ: ' + t); });
assert(out.includes('181.90') && out.includes('ร้านไทยวัสดุ') && out.includes('INV-01') && out.includes('( สมชาย )'));
assert(/<span class="cb">✓<\/span> H9 Assembly Production Line/.test(out), 'ต้องติ๊กแผนกที่เกิดค่าใช้จ่ายให้');
assert(api.mxVoucherHtml({ ...d, vendor: '<b>x</b>' }, '').includes('&lt;b&gt;x&lt;/b&gt;'), 'ชื่อร้านต้อง escape');
// ใบเปล่าไว้เขียนมือ: ไม่มีตัวเลข 0.00 โผล่
const blank = api.mxVoucherHtml(api.mxVoucherData([]), '');
assert(!blank.includes('0.00'), 'ใบเปล่าต้องไม่มียอด 0.00');

// ── ปุ่มในหน้าค่าใช้จ่ายสิ้นเปลือง ──────────────────────────────────────────
assert(htmlLf.includes('id="miscExpenseVoucherBlankBtn"'), 'ต้องมีปุ่มพิมพ์ใบเบิกเปล่า');
assert(htmlLf.includes('data-mx-voucher="'), 'แต่ละแถวต้องมีปุ่มพิมพ์ใบเบิก');
assert(htmlLf.includes("printMiscExpenseVoucher(voucherBtn.getAttribute('data-mx-voucher'))"));
assert(/rows = billId \? miscExpenseCache\.filter\(function\(r\) \{ return String\(r\.bill_id \|\| ''\) === billId; \}\) : \[item\];/.test(htmlLf),
  'พิมพ์จากแถวไหนก็ต้องได้ทั้งบิล');

// ── ต้องจบใน A4 แผ่นเดียว ─────────────────────────────────────────────────
const css = new Function(grab(/^ {4}function mxVoucherCss\(\) \{[\s\S]*?\n {4}\}/m, 'mxVoucherCss') + '\nreturn mxVoucherCss;')()();
assert(css.includes('@page{size:A4 portrait;margin:8mm}'));
assert(css.includes('#mxPageProbe{') && css.includes('height:281mm'), 'ตัววัดต้องสูงเท่าพื้นที่พิมพ์ A4 (297 - 8×2)');
const fitSrc = grab(/^ {4}function mxVoucherFitScript\(\) \{[\s\S]*?\n {4}\}/m, 'mxVoucherFitScript');
const fit = new Function(fitSrc + '\nreturn mxVoucherFitScript;')()();
assert(fit.includes('<\\/script>') === false && fit.endsWith('<\/script>'), 'ปิด script ได้ถูกต้องในหน้าต่างพิมพ์');
assert(fit.includes('beforeprint') && fit.includes('s.style.zoom='), 'รายการเยอะต้องย่อทั้งใบลงให้พอดีแผ่นเดียว');
assert(fitSrc.includes('<\\/script>'), 'ใน index.html ต้องเขียน <\\/script> ไม่งั้นปิด <script> ของหน้าหลักกลางคัน');
assert(htmlLf.includes(`mxVoucherHtml(d, appDate) + '<div id="mxPageProbe"></div>' + mxVoucherFitScript() + '</body></html>'`));
assert(htmlLf.includes('if (win.mxVoucherFit) win.mxVoucherFit(); win.print();'));

console.log('misc-expense-voucher: ok');
