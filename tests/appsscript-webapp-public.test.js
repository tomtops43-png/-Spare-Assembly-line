// web app ต้องเปิด "ทุกคนที่มีลิงก์" — เครื่องหน้างานไม่ได้ล็อกอิน Google และหน้าเว็บเรียก /exec
// แบบ credentials: 'omit' ถ้าตั้งเป็นอย่างอื่น ทุก request โดนเด้งไปหน้า login = แอปพังทั้งระบบ
const fs = require('fs');
const assert = require('assert');

const manifest = JSON.parse(fs.readFileSync('scr/appsscript.json', 'utf8'));
assert(manifest.webapp, 'appsscript.json ต้องมี webapp');
assert.strictEqual(manifest.webapp.access, 'ANYONE_ANONYMOUS', 'webapp.access ต้องเป็น ANYONE_ANONYMOUS (Anyone — ไม่ต้องล็อกอิน)');
assert.strictEqual(manifest.webapp.executeAs, 'USER_DEPLOYING', 'webapp.executeAs ต้องเป็น USER_DEPLOYING (รันด้วยสิทธิ์เจ้าของ)');

const clasp = JSON.parse(fs.readFileSync('scr/.clasp.json', 'utf8'));
assert(clasp.scriptId, 'scr/.clasp.json ต้องมี scriptId ให้ workflow deploy ใช้เป็นค่าตั้งต้น');

// workflow deploy ต้องคงขั้นตอนตรวจสิทธิ์ไว้ก่อน push
const wf = fs.readFileSync('.github/workflows/deploy-appsscript.yml', 'utf8');
const verifyAt = wf.indexOf('Verify web app stays public');
const pushAt = wf.indexOf('clasp push');
assert(verifyAt >= 0, 'workflow ต้องมีขั้นตอน Verify web app stays public');
assert(pushAt > verifyAt, 'ต้องตรวจสิทธิ์ web app ก่อน clasp push');
assert(/@google\/clasp@3\b/.test(wf), 'workflow ต้องตรึง clasp@3');

console.log('appsscript-webapp-public: ok');
