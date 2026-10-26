const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── กำลังกรอกบิลยาวๆ แล้วเผลอกด × / ยกเลิก / พื้นหลัง ข้อมูลต้องไม่หาย ─────────────
const src = [
  grab(/^ {4}function mxIsBlankRow\(row\) \{[\s\S]*?\n {4}\}/m, 'mxIsBlankRow'),
  grab(/^ {4}function mxSnapshotHasData\(snap\) \{[\s\S]*?\n {4}\}/m, 'mxSnapshotHasData')
].join('\n');
const mxSnapshotHasData = new Function(src + '\nreturn mxSnapshotHasData;')();

const blankRow = { category: '', item_name: '', qty: 1, unit: 'ชิ้น', unit_price: '' };
const empty = { date: '2026-10-26', line: '', vendor: '', receipt_no: '', receipt_url: '', remark: '', shipping: '', vat_enabled: false, vat_percent: '7', rows: [blankRow] };
assert.strictEqual(mxSnapshotHasData(empty), false, 'ฟอร์มเปล่า (มีแค่วันที่วันนี้ + VAT 7) ต้องปิดได้ทันทีไม่ถาม');
assert.strictEqual(mxSnapshotHasData(Object.assign({}, empty, { rows: [Object.assign({}, blankRow, { item_name: 'ท่อหด' })] })), true);
assert.strictEqual(mxSnapshotHasData(Object.assign({}, empty, { vendor: 'ร้านไทยวัสดุ' })), true);
assert.strictEqual(mxSnapshotHasData(Object.assign({}, empty, { line: 'H9' })), true);
assert.strictEqual(mxSnapshotHasData(null), false);

// ปุ่มปิดทุกทางต้องผ่านตัวถามยืนยัน ไม่ใช่ปิดตรงๆ
assert(htmlLf.includes("if (el) el.addEventListener('click', requestCloseMiscExpenseModal);"), '× และ ยกเลิก ต้องถามก่อนปิด');
assert(htmlLf.includes("bindBackdropClose(modal, requestCloseMiscExpenseModal)"), 'กดพื้นหลังต้องถามก่อนปิด');
assert(/function requestCloseMiscExpenseModal\(\) \{[\s\S]{0,600}showConfirmDialog\(/.test(htmlLf));
// ร่างต้องถูกเก็บระหว่างพิมพ์ ดึงกลับตอนเปิดใหม่ และล้างทิ้งเมื่อบันทึกสำเร็จ
assert(htmlLf.includes("form.addEventListener('input', mxSaveDraft);"));
assert(/var draft = mxLoadDraft\(\);\s*mxFillNewForm\(draft\);/.test(htmlLf), 'เปิดบิลใหม่ต้องดึงร่างกลับมา');
assert(/if \(!expenseId\) mxClearDraft\(\);\s*closeMiscExpenseModal\(\);/.test(htmlLf), 'บันทึกสำเร็จต้องล้างร่าง');
assert(htmlLf.includes('id="miscExpenseClearDraft"'), 'ต้องมีปุ่มล้างฟอร์มไว้ทิ้งร่างเอง');

console.log('Misc expense draft guard checks passed');
