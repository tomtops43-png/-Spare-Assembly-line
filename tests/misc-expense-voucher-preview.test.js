// ปุ่ม "🧾 ดูใบเบิก" ในฟอร์มค่าใช้จ่าย — ดูหน้าตาใบเบิกจ่ายจากข้อมูลที่กรอกอยู่ก่อนกดบันทึก
const { html, grabFn, assert } = require('./_export-extract');

assert(html.includes('<button id="miscExpensePreviewVoucher" type="button"'), 'มีปุ่มดูใบเบิกในฟอร์ม');
assert(html.indexOf('id="miscExpensePreviewVoucher"') < html.indexOf('id="miscExpenseSubmit"'), 'ปุ่มอยู่ข้างปุ่มบันทึก');
assert(html.includes("previewBtn.addEventListener('click', previewMiscExpenseVoucher)"));

const fields = { miscExpenseDate: '2026-10-29', miscExpenseLine: 'H9', miscExpenseVendor: 'ฮาร์ดแวร์เฮาส์', miscExpenseReceiptNo: 'CS4691029-0433', miscExpenseRemark: '' };
function run(items) {
  const calls = { print: null, msg: null };
  const fn = new Function('mxCollectItems', 'mxEl', 'currentUser', 'printMiscExpenseVoucher', 'setMiscExpenseFormMessage',
    grabFn('previewMiscExpenseVoucher') + '\nreturn previewMiscExpenseVoucher;')(
    function() { return items; },
    function(id) { return { value: fields[id] }; },
    { username: 'Nun' },
    function(id, rows) { calls.print = { id: id, rows: rows }; },
    function(m) { calls.msg = m; });
  fn();
  return calls;
}

// ใช้รายการเดียวกับที่จะบันทึก (mxCollectItems) + ข้อมูลหัวบิล → ใบออกมาเหมือนตอนพิมพ์หลังบันทึก
const r = run([{ category: 'เครื่องมือช่าง', item_name: 'คีม', qty: 1, unit: 'ชิ้น', total_amount: 105 }]);
assert.strictEqual(r.print.id, '', 'ยังไม่มีเลขรายการ เพราะยังไม่บันทึก');
assert.deepStrictEqual(r.print.rows[0], { date: '2026-10-29', line: 'H9', vendor: 'ฮาร์ดแวร์เฮาส์', receipt_no: 'CS4691029-0433', remark: '', created_by: 'Nun', category: 'เครื่องมือช่าง', item_name: 'คีม', qty: 1, unit: 'ชิ้น', total_amount: 105 });
// ยังไม่กรอกรายการ → เตือน ไม่เปิดใบเปล่า
const empty = run([]);
assert.strictEqual(empty.print, null);
assert(empty.msg.includes('ยังไม่มีรายการ'));

// พิมพ์จากฟอร์ม: ใช้แถวที่ส่งมาแทนการค้นจากรายการที่บันทึกแล้ว
assert(grabFn('printMiscExpenseVoucher').includes('var rows = Array.isArray(previewRows) ? previewRows : [];'));

console.log('misc-expense-voucher-preview: OK');
