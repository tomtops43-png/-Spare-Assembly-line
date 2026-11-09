// กดเลือกไลน์ใน Sidebar ระหว่างอยู่หน้าเบิกของสิ้นเปลือง → ต้องพาไปหน้า Stock ของไลน์นั้น (ไม่ค้างหน้าเดิม)
const { grabFn, assert } = require('./_export-extract');

const src = grabFn('renderLineTabs');
assert(/if \(csIsPageVisible\(\)\) switchToStockPage\(\);/.test(src), 'คลิกไลน์ตอนอยู่หน้าของสิ้นเปลืองต้องสลับไปหน้า Stock');
// ต้องสลับหลังตั้งไลน์ใหม่แล้ว (หน้า Stock จะได้แสดงของไลน์ที่เพิ่งเลือก)
assert(src.indexOf('setCurrentLine(') < src.indexOf('switchToStockPage()'));

console.log('consumable-line-switch: OK');
