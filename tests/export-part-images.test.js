// รูปอะไหล่ที่ฝังลงไฟล์ Excel
// ผู้จัดการเอาไฟล์ Excel ไปใช้ต่อ (กรอง/pivot) แต่ขอเห็นรูปของด้วย จึงทำเป็นชีตแยก
// ไม่แปะรูปในชีต Master เพราะทุกแถวจะสูง 50px แล้วใช้เป็นตารางข้อมูลไม่สะดวก
//
// จุดที่กันไว้:
//  1) รูปมาจาก Google Drive — เบราว์เซอร์อาจอ่านไบต์ไม่ได้เพราะ CORS ต้องมีทางถอย
//  2) รูปต้องถูกย่อ ห้ามฝังไฟล์ต้นฉบับหลาย MB ลงไป (ไฟล์จะเปิดไม่ไหว)
//  3) รูปต้องคงสัดส่วน ไม่ยืด (รูปอะไหล่ที่ยืดแล้วดูไม่ออกว่าชิ้นไหน)
//  4) ดึงรูปไม่ได้ต้องไม่ทำให้ไฟล์ข้อมูลออกไม่ได้ — รูปเป็นของแถม
//  5) แถวกับรูปต้องตรงกัน ไม่ใช่รูปอะไหล่ A ไปอยู่แถวอะไหล่ B
const { html, backend, grabFn, grabVar, grabScalar, buildModule, baseHelpers, assert } = require('./_export-extract');

// ── ย่อรูป: ต้องคงสัดส่วนและไม่เกินกล่องที่กำหนด ──────────────────────
const shrinkMod = (function() {
  // stub canvas เท่าที่ xpShrinkImage ใช้ — node ไม่มี DOM
  const calls = [];
  const doc = {
    createElement: function() {
      const c = { width: 0, height: 0 };
      c.getContext = function() {
        return {
          fillRect: function() {},
          drawImage: function(img, x, y, w, h) { calls.push({ w: w, h: h }); },
          set fillStyle(v) {}
        };
      };
      c.toDataURL = function(type, q) { return 'data:image/jpeg;base64,STUB(' + c.width + 'x' + c.height + ',q=' + q + ')'; };
      return c;
    }
  };
  const src = [grabScalar('XP_IMAGE_BOX'), grabFn('xpShrinkImage'), 'return { xpShrinkImage: xpShrinkImage, XP_IMAGE_BOX: XP_IMAGE_BOX, calls: calls };'].join('\n');
  return new Function('document', 'calls', src)(doc, calls);
})();

const BOX = shrinkMod.XP_IMAGE_BOX;
assert(BOX > 0 && BOX <= 200, 'กล่องรูปควรเล็ก (<=200px) ให้ไฟล์ไม่บวม แต่ตั้งไว้ ' + BOX);

// รูปนอน 800x600 → ต้องย่อให้ด้านยาวเท่ากล่อง และสัดส่วนคงเดิม
const wide = shrinkMod.xpShrinkImage({ naturalWidth: 800, naturalHeight: 600 });
assert.strictEqual(wide.width, BOX, 'ด้านยาวต้องเท่ากล่องพอดี');
assert.strictEqual(wide.height, Math.round(600 * (BOX / 800)), 'ด้านสั้นต้องย่อตามสัดส่วน');
assert(Math.abs((wide.width / wide.height) - (800 / 600)) < 0.02, 'สัดส่วนต้องไม่เพี้ยน');
// รูปตั้ง 600x900
const tall = shrinkMod.xpShrinkImage({ naturalWidth: 600, naturalHeight: 900 });
assert.strictEqual(tall.height, BOX, 'รูปตั้ง: ด้านยาวคือความสูง ต้องเท่ากล่อง');
assert(tall.width < tall.height, 'รูปตั้งต้องยังตั้งอยู่');
// รูปที่เล็กกว่ากล่องอยู่แล้ว ต้องไม่ถูกขยาย (ขยายแล้วเบลอเปล่าๆ และไฟล์ใหญ่ขึ้น)
const small = shrinkMod.xpShrinkImage({ naturalWidth: 40, naturalHeight: 30 });
assert.strictEqual(small.width, 40, 'รูปเล็กกว่ากล่องต้องไม่ถูกขยาย');
assert.strictEqual(small.height, 30);
// รูปที่โหลดไม่สำเร็จ (ขนาด 0) ต้องคืน null ไม่ใช่รูปเปล่า
assert.strictEqual(shrinkMod.xpShrinkImage({ naturalWidth: 0, naturalHeight: 0 }), null, 'รูปขนาด 0 ต้องคืน null');
// ต้องออกเป็น JPEG คุณภาพต่ำกว่า 1 (ไม่ใช่ PNG ที่ใหญ่กว่าหลายเท่า)
assert(/image\/jpeg/.test(wide.dataUrl), 'ต้องบีบเป็น JPEG ให้ไฟล์เล็ก');
assert(/q=0\.\d/.test(wide.dataUrl), 'ต้องตั้งคุณภาพต่ำกว่า 1');

// ── ชีตรูปอะไหล่: แถวกับรูปต้องตรงกัน ───────────────────────────────
const sheetMod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", category: "", txnType: "", groupBy: "month", lang: "th", includeImages: true };',
    grabFn('extractDriveFileId'),
    grabFn('xpBuildPartPhotoSheet')
  ]),
  '{ xpBuildPartPhotoSheet: xpBuildPartPhotoSheet, setLang: function(l) { xpFilters.lang = l; } }'
);

const items = [
  { __sheet: 'H9', no: '1', name: 'สายพาน', model: 'B-77', line: 'H9', category: 'กลไก', location: 'B-1', stock: 6, unit: 'PCS', image_main_url: 'https://drive.google.com/thumbnail?id=AAAAAAAAAAAAAAAAAAAAA1&sz=w1200' },
  { __sheet: 'H9', no: '2', name: 'ไม่มีรูป', model: 'NP-1', line: 'H9', category: 'ไฟฟ้า', location: 'B-2', stock: 3, unit: 'PCS', image_main_url: '' },
  { __sheet: 'LugScrew', no: '3', name: 'เบรกเกอร์', model: 'BK-10', line: 'Lug&Screw', category: 'ไฟฟ้า', location: 'A-1', stock: 4, unit: 'PCS', image_main_url: 'https://drive.google.com/thumbnail?id=BBBBBBBBBBBBBBBBBBBBB2&sz=w1200' },
  { __sheet: 'LugScrew', no: '4', name: 'ดึงรูปไม่ได้', model: 'FA-1', line: 'Lug&Screw', category: 'อื่นๆ', location: 'A-2', stock: 1, unit: 'PCS', image_main_url: 'https://drive.google.com/thumbnail?id=CCCCCCCCCCCCCCCCCCCCC3&sz=w1200' }
];
const imgRes = {
  byId: {
    'AAAAAAAAAAAAAAAAAAAAA1': { dataUrl: 'data:image/jpeg;base64,AAA', width: 96, height: 72 },
    'BBBBBBBBBBBBBBBBBBBBB2': { dataUrl: 'data:image/jpeg;base64,BBB', width: 60, height: 96 }
    // CCC ดึงไม่ได้ → ต้องไม่มีแถวในชีต
  },
  method: 'client', failed: 1, capped: false, total: 3
};
const photo = sheetMod.xpBuildPartPhotoSheet(items, imgRes);

assert.strictEqual(photo.table.rows.length, 2, 'ต้องมีเฉพาะรายการที่ได้รูปจริง (ไม่มีรูป/ดึงไม่ได้ ต้องไม่มีแถวเปล่า)');
assert.strictEqual(photo.rowImages.length, 2, 'จำนวนรูปต้องเท่าจำนวนแถว');
const nameIdx = photo.table.headers.indexOf('ชื่อรายการ');
assert.strictEqual(photo.table.rows[0][nameIdx], 'สายพาน');
assert.strictEqual(photo.table.rows[1][nameIdx], 'เบรกเกอร์');
// รูปต้องผูกกับแถวที่ถูกต้อง — สลับกันคือรูปผิดชิ้น
assert.strictEqual(photo.rowImages[0].rowIndex, 0);
assert.strictEqual(photo.rowImages[0].dataUrl, 'data:image/jpeg;base64,AAA', 'แถวแรก (สายพาน) ต้องได้รูปของสายพาน');
assert.strictEqual(photo.rowImages[1].rowIndex, 1);
assert.strictEqual(photo.rowImages[1].dataUrl, 'data:image/jpeg;base64,BBB', 'แถวสอง (เบรกเกอร์) ต้องได้รูปของเบรกเกอร์');
// ต้องมีคอลัมน์อ้างอิงพอจับคู่กลับไปที่ชีต Master ได้
['รูป', 'ชีตต้นทาง', 'NO', 'ชื่อรายการ', 'รุ่น / Part No.', 'ไลน์', 'คงเหลือ', 'ลิงก์รูป'].forEach(function(h) {
  assert(photo.table.headers.indexOf(h) > -1, 'ชีตรูปต้องมีคอลัมน์ ' + h);
});
assert.strictEqual(photo.table.headers[0], 'รูป', 'คอลัมน์รูปต้องเป็นคอลัมน์แรก (ที่วางรูป)');
assert.strictEqual(photo.table.rows[0][0], '', 'ช่องคอลัมน์รูปต้องว่างไว้ให้รูปทับ');
// โหมดอังกฤษ
sheetMod.setLang('en');
const photoEn = sheetMod.xpBuildPartPhotoSheet(items, imgRes);
assert.strictEqual(photoEn.table.headers[0], 'Photo', 'หัวคอลัมน์รูปต้องถูกแปล');
assert(photoEn.table.headers.indexOf('Photo Link') > -1, 'คอลัมน์ลิงก์รูปต้องถูกแปล');
sheetMod.setLang('th');

// รูปเดียวกันใช้ซ้ำหลายรายการได้ — ต้องดึงครั้งเดียวแต่แสดงทุกแถว
const shared = [
  { __sheet: 'H9', no: '9', name: 'ชิ้น A', model: 'X', image_main_url: 'https://drive.google.com/thumbnail?id=AAAAAAAAAAAAAAAAAAAAA1' },
  { __sheet: 'H9', no: '10', name: 'ชิ้น B', model: 'Y', image_main_url: 'https://drive.google.com/thumbnail?id=AAAAAAAAAAAAAAAAAAAAA1' }
];
const sharedSheet = sheetMod.xpBuildPartPhotoSheet(shared, imgRes);
assert.strictEqual(sharedSheet.table.rows.length, 2, 'รูปเดียวกันใช้กับหลายรายการต้องได้ครบทุกแถว');

// ── ทางถอยเมื่อเบราว์เซอร์อ่านไบต์รูปไม่ได้ (CORS) ────────────────────
const loadSrc = grabFn('xpLoadImageViaCanvas');
assert(/img\.crossOrigin = 'anonymous'/.test(loadSrc), "ต้องขอรูปแบบ crossOrigin='anonymous' ไม่งั้น canvas อ่านไบต์ไม่ได้เลย");
assert(/err\.xpTainted = true/.test(loadSrc), 'ต้องแยกได้ว่า error เป็นเพราะ CORS (tainted canvas) หรือโหลดรูปไม่สำเร็จ');
assert(/setTimeout/.test(loadSrc), 'ต้องมี timeout — รูปที่ค้าง pending จะทำให้งานไม่จบ');

const fetchSrc = grabFn('xpFetchPartImages');
assert(/if \(err && err\.xpTainted\) \{ tainted = true; return; \}/.test(fetchSrc), 'เจอ tainted ต้องสลับไปทางเซิร์ฟเวอร์');
assert(/xpLoadImagesViaServer/.test(fetchSrc), 'ต้องมีทางถอยไปขอรูปจากเซิร์ฟเวอร์');
assert(/if \(!tainted\)/.test(fetchSrc), 'ถ้าอ่านฝั่งเว็บได้ต้องไม่ยิงเซิร์ฟเวอร์เลย (เร็วกว่าและไม่กิน quota)');
assert(/XP_IMAGE_MAX/.test(fetchSrc), 'ต้องมีเพดานจำนวนรูปต่อไฟล์');
assert(/t\.id === id/.test(fetchSrc), 'รูปเดียวกันต้องดึงครั้งเดียว');
// เพดานต้องสมเหตุสมผล
const MAX = buildModule([grabScalar('XP_IMAGE_MAX')], 'XP_IMAGE_MAX');
assert(MAX >= 100 && MAX <= 1000, 'เพดานรูปควรอยู่ 100-1000 แต่ตั้งไว้ ' + MAX);
const CONC = buildModule([grabScalar('XP_IMAGE_CONCURRENCY')], 'XP_IMAGE_CONCURRENCY');
assert(CONC >= 2 && CONC <= 12, 'จำนวนที่โหลดพร้อมกันควรอยู่ 2-12 แต่ตั้งไว้ ' + CONC);

// ก้อนที่ส่งให้เซิร์ฟเวอร์ต้องไม่เกินที่ฝั่งเซิร์ฟเวอร์รับ
const CHUNK = buildModule([grabScalar('XP_IMAGE_SERVER_CHUNK')], 'XP_IMAGE_SERVER_CHUNK');
const SRV_MAX = buildModule([grabScalar('EXPORT_IMAGE_MAX_IDS', backend)], 'EXPORT_IMAGE_MAX_IDS');
assert(CHUNK <= SRV_MAX, 'ก้อนที่ฝั่งเว็บส่ง (' + CHUNK + ') ต้องไม่เกินที่ backend รับ (' + SRV_MAX + ') ไม่งั้นรูปหายเงียบๆ');
const srvSrc = grabFn('xpLoadImagesViaServer');
assert(/\.catch\(/.test(srvSrc), 'ก้อนไหนพังต้องข้ามไป รูปที่เหลือยังได้');

// ── ฝั่งเซิร์ฟเวอร์ ──────────────────────────────────────────────────
const beSrc = grabFn('exportPartImages', backend, 0);
assert(/requireExportAccess/.test(beSrc), 'ต้องตรวจสิทธิ์ก่อนอ่านไฟล์จาก Drive');
assert(/\/\^\[A-Za-z0-9_-\]\{20,\}\$\//.test(beSrc), 'ต้องรับเฉพาะรูปแบบ Drive file id — กันการยัด URL อื่นให้เซิร์ฟเวอร์ไปดึงแทน');
assert(/EXPORT_IMAGE_MAX_IDS/.test(beSrc), 'ต้องจำกัดจำนวนไอดีต่อคำขอ (Apps Script มีเพดานเวลา 6 นาที)');
const thumbSrc = grabFn('fetchDriveThumbnailBase64', backend, 0);
assert(/getThumbnail/.test(thumbSrc), 'ควรลอง getThumbnail() ก่อน (ไม่ต้องยิง HTTP)');
assert(/UrlFetchApp\.fetch/.test(thumbSrc), 'ถ้าไม่มีรูปย่อ ต้องขอผ่าน endpoint thumbnail ที่ระบุขนาดได้');
assert(/ScriptApp\.getOAuthToken\(\)/.test(thumbSrc), 'ต้องแนบ token ไม่งั้นอ่านไฟล์ส่วนตัวไม่ได้');
assert(!/file\.getBlob\(\)/.test(thumbSrc), 'ห้ามใช้ getBlob() ของไฟล์ต้นฉบับ — รูปหลาย MB ฝังลง Excel แล้วไฟล์บวมจนเปิดไม่ไหว');
assert(/indexOf\('image\/'\) !== 0/.test(thumbSrc), 'ต้องเช็คว่าเป็นไฟล์รูปจริง ไม่ใช่ฝังไฟล์อะไรก็ได้ลง Excel');
// route ทั้งสองทาง
assert(/if \(action === 'exportPartImages'\) return respond\(exportPartImages\(e\.parameter\), e\);/.test(backend), 'ขาด route ใน doGet');
assert(/if \(action === 'exportPartImages'\) return respond\(exportPartImages\(body\), e\);/.test(backend), 'ขาด route ใน doPost');

// ── ExcelJS: วางรูปต่อแถวได้ถูกต้อง ──────────────────────────────────
const writeSrc = grabFn('xpWriteXlsxStyled');
assert(/spec\.rowImages \|\| \[\]/.test(writeSrc), 'ต้องรองรับรูปต่อแถว และชีตที่ไม่มีรูปต้องทำงานเหมือนเดิม');
assert(/ws\.getRow\(rh\)\.height/.test(writeSrc), 'ต้องขยายความสูงแถวก่อน ไม่งั้นรูปล้นทับแถวถัดไป');
assert(/ws\.getColumn\(1\)\.width/.test(writeSrc), 'ต้องขยายความกว้างคอลัมน์แรกให้พอดีรูป');
assert(/Math\.min\(cellPx \/ w, cellPx \/ h, 1\)/.test(writeSrc), 'ต้องคงสัดส่วนรูปในกรอบ ไม่ยืดรูป');
assert(/extension: 'jpeg'/.test(writeSrc), 'รูปอะไหล่ถูกบีบเป็น JPEG ต้องประกาศชนิดให้ตรง');
assert(/im\.rowIndex \+ 1/.test(writeSrc), 'ต้องเลื่อนตำแหน่งรูปลง 1 แถวเพราะแถวแรกเป็นหัวตาราง');

// ── checkbox แนบรูป ─────────────────────────────────────────────────
assert(html.indexOf('id="xpIncludeImages"') > -1, 'ต้องมี checkbox แนบรูปในแถบตัวกรอง');
assert(/id="xpIncludeImages" type="checkbox" checked/.test(html), 'ควรเปิดไว้เป็นค่าเริ่มต้น (ผู้ใช้ขอรูปมาเอง)');
const readSrc = grabFn('xpReadFilterInputs');
assert(/xpFilters\.includeImages = !im \|\| !!im\.checked;/.test(readSrc), 'ต้องอ่านค่า checkbox');
const appendSrc = grabFn('xpAppendPhotoSheet');
assert(/if \(!xpFilters\.includeImages\) return Promise\.resolve\(null\);/.test(appendSrc), 'ติ๊กออกแล้วต้องไม่ดึงรูปเลย (ไม่ใช่ดึงแล้วไม่ใช้)');
assert(/\.catch\(function\(err\)/.test(appendSrc), 'ดึงรูปพลาดต้องไม่ทำให้ไฟล์ข้อมูลออกไม่ได้ — รูปเป็นของแถม');
assert(/ไฟล์ออกครบทุกชีตแต่ไม่มีชีตรูป/.test(appendSrc), 'ต้องบอกผู้ใช้ตรงๆ ว่าไฟล์ยังครบ แค่ไม่มีรูป');
const summarySrc = grabFn('xpRenderFilterSummary');
assert(/includeImages/.test(summarySrc), 'บรรทัดสรุปต้องบอกว่าจะแนบรูปหรือไม่');

// ── ลำดับชีตและพจนานุกรม ─────────────────────────────────────────────
const wbSrc = grabFn('xpRunWorkbook');
assert(wbSrc.indexOf('xpAppendPhotoSheet') < wbSrc.indexOf("xpT('99 พจนานุกรมข้อมูล')"),
  'ต้องแนบชีตรูปก่อนสร้างพจนานุกรม ไม่งั้นคอลัมน์ของชีตรูปไม่มีคำอธิบาย');
assert(wbSrc.indexOf('xpAppendPhotoSheet') < wbSrc.indexOf("xpT('98 ลิงก์ไฟล์แนบ')"),
  'ชีตรูป (11) ต้องมาก่อนชีต 98/99 ให้ลำดับชีตในไฟล์เรียงถูก');
const dsExportSrc = grabFn('xpRunSelectedExport');
assert(/xpAppendPhotoSheet/.test(dsExportSrc), 'ระดับ 2 ที่ออกเป็น xlsx ก็ควรแนบรูปได้');
assert(/keys\.indexOf\('stockMaster'\) > -1/.test(dsExportSrc), 'แนบรูปเฉพาะเมื่อเลือกชุด Master อะไหล่ (รูปผูกกับรายการอะไหล่)');

// ── คำอธิบายคอลัมน์รูปต้องอยู่ในพจนานุกรมข้อมูล ──────────────────────
const XP_DICTIONARY = buildModule([grabVar('XP_DICTIONARY')], 'XP_DICTIONARY');
assert(XP_DICTIONARY['รูป'], 'คอลัมน์รูปต้องมีคำอธิบายในพจนานุกรมข้อมูล');
assert(XP_DICTIONARY['ลิงก์รูป'], 'คอลัมน์ลิงก์รูปต้องมีคำอธิบาย');
assert(/ย่อ/.test(XP_DICTIONARY['รูป']), 'ต้องบอกว่าเป็นรูปย่อ ไม่ใช่รูปต้นฉบับ');

console.log('export-part-images: OK (กล่องรูป ' + BOX + 'px, เพดาน ' + MAX + ' รูป, ก้อนละ ' + CHUNK + '/' + SRV_MAX + ')');
