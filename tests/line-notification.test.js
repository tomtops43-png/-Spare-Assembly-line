const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const backendLf = backend.replace(/\r\n/g, '\n');

// ── ต้องใช้ Messaging API ไม่ใช่ LINE Notify (LINE Notify ปิดบริการไปแล้ว) ────
assert(backend.includes('https://api.line.me/v2/bot/message/push'), 'ต้องยิง push ผ่าน Messaging API');
assert(!/notify-api\.line\.me/.test(backend), 'ห้ามใช้ LINE Notify ที่ปิดบริการไปแล้ว');

// ── token ห้ามหลุดออกจากฝั่งเซิร์ฟเวอร์ ──────────────────────────────────────
assert(/PropertiesService\.getScriptProperties\(\)/.test(backend) && backend.includes("getProperty('LINE_CHANNEL_TOKEN')"));
const statusSrc = backendLf.match(/function getLineStatus\(payload\) \{[\s\S]*?\n\}/);
assert(statusSrc, 'ต้องมี getLineStatus');
assert(!/return[\s\S]*cfg\.token[^s]/.test(statusSrc[0].replace(/!!\(cfg\.token && cfg\.target\)/g, '').replace(/!!cfg\.token/g, '')),
  'getLineStatus ต้องไม่คืนค่า token ออกไป บอกได้แค่ว่าตั้งค่าครบหรือยัง');
assert(/configured: !!\(cfg\.token && cfg\.target\)/.test(backend));
assert(/function sendLineTestMessage\(payload\) \{[\s\S]{0,120}requireAdminUser/.test(backend), 'ปุ่มทดสอบต้องเป็น Admin เท่านั้น');
assert(/function getLineStatus\(payload\) \{[\s\S]{0,120}requireAdminUser/.test(backend));

// ── แจ้งเตือนพังห้ามทำให้งานหลักล้มตาม ───────────────────────────────────────
// สร้าง PR สำเร็จแล้ว แต่ LINE ล่ม/token หมดอายุ ต้องไม่กลายเป็น "สร้าง PR ไม่สำเร็จ"
assert(/function sendLineMessage\(text\) \{[\s\S]{0,1600}muteHttpExceptions: true/.test(backend), 'ต้อง mute เพื่อไม่ให้ HTTP error กลายเป็น exception');
assert(/function sendLineMessage\(text\) \{[\s\S]{0,1900}catch \(err\) \{[\s\S]{0,160}return \{ ok: false, reason: 'FETCH_ERROR'/.test(backend));
assert(/function notifyLine\(text\) \{[\s\S]{0,700}catch \(err\) \{[\s\S]{0,200}Logger\.log\('notifyLine error/.test(backend));
// ยังไม่ตั้งค่า = เงียบ ไม่ใช่ spam log ทุกเช้า
assert(/if \(!result\.ok && result\.reason !== 'NOT_CONFIGURED'\)/.test(backend));
assert(/if \(isLineConfigured\(\)\) results\.line = notifyLine\(buildDailyLineDigest\(results\)\);/.test(backend));
assert(/try \{[\s\S]{0,200}buildDailyLineDigest\(results\)[\s\S]{0,200}\} catch \(e3\)/.test(backend), 'digest พังต้องไม่ทำให้งานเช้าทั้งชุดพัง');

// ── ข้อความยาวเกินลิมิต LINE ต้องถูกตัด ไม่ใช่ให้ API ตีกลับทั้งก้อน ──────────
assert(/if \(body\.length > 4900\) body = body\.slice\(0, 4900\)/.test(backend));

// ── เนื้อหาที่ต้องแจ้ง ──────────────────────────────────────────────────────
const digestSrc = backendLf.match(/function buildDailyLineDigest\(jobResults\) \{[\s\S]*?\n\}/);
assert(digestSrc, 'ต้องมี buildDailyLineDigest');
['หมดสต็อก', 'ต่ำกว่า Min', 'PR รออนุมัติ', 'Auto-PR', 'เบิกผิดปกติ'].forEach(function(k) {
  assert(digestSrc[0].indexOf(k) > -1, 'สรุปประจำวันต้องมี "' + k + '"');
});
// ตัวเลขเฉยๆ ไม่พอให้ตัดสินใจ ต้องบอกชื่อของที่หมดด้วย
assert(/outItems\.slice\(0, 5\)\.forEach/.test(backend), 'ต้องโชว์ชื่อของที่หมดจริงๆ ไม่ใช่แค่จำนวน');
assert(/if \(outItems\.length > 5\) lines\.push/.test(backend), 'ของหมดเยอะต้องตัดแล้วบอกว่าเหลืออีกกี่รายการ');
// ใช้ผลของงานเช้าที่คำนวณไปแล้ว ไม่คำนวณซ้ำให้เปลือง quota
assert(/function buildDailyLineDigest\(jobResults\)/.test(backend) && /buildDailyLineDigest\(results\)/.test(backend));
// PR ใหม่ต้องรู้ทันที ไม่ใช่รอสรุปเช้าวันถัดไป
assert(/if \(isLineConfigured\(\)\) \{[\s\S]{0,400}'📋 PR ใหม่รออนุมัติ'/.test(backend));
assert(/return \{ status: 'success', pr_id: prId, pr_status: 'PENDING'/.test(backend), 'ต้องยังคืนผลสร้าง PR ตามเดิม');

// ── routing ────────────────────────────────────────────────────────────────
["getLineStatus", "sendLineTestMessage"].forEach(function(a) {
  assert(backend.includes("if (action === '" + a + "') return respond(" + a + "(e.parameter), e);"), a + ' ต้องเรียกผ่าน GET ได้');
  assert(backend.includes("if (action === '" + a + "') return respond(" + a + "(body), e);"), a + ' ต้องเรียกผ่าน POST ได้');
});

// ── หน้า Admin ─────────────────────────────────────────────────────────────
assert(htmlLf.includes('id="lineNotifyStatusBadge"') && htmlLf.includes('id="lineNotifyTestBtn"'));
assert(htmlLf.includes('🔔 แจ้งเตือนเข้า LINE'));
assert(/setTabStyles\('admin'\);[\s\S]{0,120}loadLineNotifyStatus\(\);/.test(htmlLf), 'เข้าหน้า Admin ต้องเช็คสถานะให้เลย');
// ยังไม่ตั้งค่า = ปุ่มทดสอบกดไม่ได้ (กดแล้วขึ้น error เปล่าๆ)
assert(/if \(testBtn\) testBtn\.disabled = !configured;/.test(htmlLf));
assert(htmlLf.includes('LINE_CHANNEL_TOKEN') && htmlLf.includes('LINE_TARGET_ID'), 'ต้องมีวิธีตั้งค่าให้อ่านในหน้าเว็บ');
assert(htmlLf.includes('LINE Notify ปิดบริการไปแล้ว'), 'ต้องเตือนไม่ให้ไปตั้ง LINE Notify ผิดตัว');

console.log('LINE notification checks passed');
