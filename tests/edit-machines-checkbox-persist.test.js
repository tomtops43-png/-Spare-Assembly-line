const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

// ── หัวใจของบั๊ก: normalizeRecord สร้าง object ใหม่แบบ whitelist ────────────────
// ถ้าไม่ map machines ต่อ ค่าที่ติ๊กไว้จะถูกทิ้งทุกครั้งที่โหลดรายการ → เปิดหน้าแก้ไข
// ช่องติ๊กจะว่างทั้งที่ในชีทมีค่า แล้วพอกดบันทึกก็เขียนค่าว่างทับของเดิม
const normalizeSrc = script.match(/function normalizeRecord\(item, index\)[\s\S]*?\n    \}/)[0];
assert(/machines:\s*firstDefined\(/.test(normalizeSrc),
  'normalizeRecord ต้อง map ฟิลด์ machines ต่อ ไม่งั้นเครื่องจักรที่ผูกไว้จะหาย');

// ฟิลด์อื่นที่หน้าแก้ไขใช้ ต้องอยู่ใน normalizeRecord ด้วย (กันหล่นแบบเดียวกันอีก)
['coil_size', 'brand', 'category', 'drawing_url'].forEach(function(field) {
  assert(normalizeSrc.indexOf(field + ': firstDefined(') > -1, 'normalizeRecord ต้องมีฟิลด์ ' + field);
});

// ── เปิดหน้าแก้ไข ต้องส่งค่าที่บันทึกไว้ไปติ๊กให้ ──────────────────────────────
assert(/renderEditMachinesCheckboxes\(item\.line \|\| '', String\(item\.machines \|\| ''\)\.split\(','\)/.test(script),
  'openEditModal ต้องส่ง item.machines ไปให้ renderEditMachinesCheckboxes');

// ── กันเซฟทับตอนรายการเครื่องจักรยังโหลดไม่เสร็จ ─────────────────────────────
const renderSrc = script.match(/function renderEditMachinesCheckboxes\(line, selectedNames\)[\s\S]*?\n    \}/)[0];
assert(/data-machines-ready', '0'/.test(renderSrc), 'ต้องรีเซ็ตธง ready ก่อนเริ่มโหลด');
// ทุกทางที่วาด checkbox เสร็จต้องตั้งธง ready='1' — ตอนนี้มี 3 ทาง:
// (1) ไลน์นี้ไม่มีเครื่องจักร (2) รายการเรียบ ไม่มีขนาดในชื่อ (3) แบบจัดกลุ่มตามขนาด
// ถ้าเพิ่มทางใหม่แล้วลืมตั้งธง = กดเซฟแล้วเครื่องจักรที่ผูกไว้จะถูกลบทิ้ง
assert((renderSrc.match(/data-machines-ready', '1'/g) || []).length === 3,
  "ต้องตั้งธง ready='1' ครบทุกทางที่วาด checkbox เสร็จ (ไม่มีเครื่อง / รายการเรียบ / จัดกลุ่ม)");
assert(/data-machines-pending/.test(renderSrc), 'ต้องจำค่าที่ขอให้ติ๊กไว้ เผื่อเปลี่ยน Line ระหว่างโหลด');

assert(/machines: isEditMachinesReady\(\) \? getSelectedEditMachines\(\)\.join\(', '\) : String\(oldItem\.machines \|\| ''\)/.test(script),
  'ตอนบันทึก ถ้ารายการยังไม่พร้อมต้องคงค่าเดิมไว้ ห้ามส่งค่าว่างไปทับ');

// เปลี่ยน Line ระหว่างที่ยังโหลดไม่เสร็จ ต้องไม่ทำให้ค่าที่ติ๊กไว้หลุด
assert(/renderEditMachinesCheckboxes\(eLineInputForMachines\.value\.trim\(\), getEditMachinesSelection\(\)\)/.test(script),
  'handler ของช่อง Line ต้องใช้ getEditMachinesSelection() ไม่ใช่อ่าน checkbox ตรงๆ');

// ── backend ต้องส่ง machines กลับมาในโหมด lite ด้วย ───────────────────────────
// (หน้ารายการโหลดแบบ lite ถ้าตัดฟิลด์นี้ทิ้ง ช่องติ๊กก็ว่างเหมือนเดิม)
const machinesRead = backend.match(/machines: pickRowValue\(row, map, \[[^\]]*\], ''\)/);
assert(machinesRead, 'backend ต้องอ่านคอลัมน์ Machines ส่งกลับมา');
assert(!/machines: isLiteRead \?/.test(backend), 'ห้ามตัด machines ทิ้งในโหมด lite');

console.log('Edit machines checkbox persistence checks passed');
