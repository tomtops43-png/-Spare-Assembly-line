const fs = require('fs');
const assert = require('assert');

// หน้าเว็บ (GitHub Pages) เรียก backend แบบไม่ login — webapp ต้องเปิดให้ทุกคนและรันในนามผู้ deploy
// workflow deploy-appsscript.yml เช็กซ้ำอีกชั้นก่อน deploy แต่เทสต์นี้จับได้ตั้งแต่ตอน PR
const manifest = JSON.parse(fs.readFileSync('appsscript.json', 'utf8'));
assert.strictEqual(manifest.webapp.access, 'ANYONE_ANONYMOUS', 'webapp.access = ANYONE_ANONYMOUS');
assert.strictEqual(manifest.webapp.executeAs, 'USER_DEPLOYING', 'webapp.executeAs = USER_DEPLOYING');
assert.strictEqual(manifest.timeZone, 'Asia/Bangkok', 'timezone ของโรงงาน');

// .clasp.json ใน repo เป็นค่าหลอก — scriptId จริงมาจาก secret SCRIPT_ID ตอน deploy
const clasp = JSON.parse(fs.readFileSync('.clasp.json', 'utf8'));
assert.strictEqual(clasp.scriptId, 'YOUR_ACTUAL_SCRIPT_ID', '.clasp.json ใช้ scriptId หลอก');
assert.strictEqual(clasp.rootDir, 'src', 'clasp push จาก src/');

// โค้ด backend อยู่ใน src/ และ manifest ตัวจริงอยู่ที่ root เท่านั้น (src/appsscript.json ถูก copy ตอน deploy)
assert(fs.readdirSync('src').some(function(f) { return f.endsWith('.gs'); }), 'มีไฟล์ .gs ใน src/');
assert(!fs.existsSync('scr'), 'โฟลเดอร์ scr/ เดิมย้ายไป src/ แล้ว');
assert(/^src\/appsscript\.json$/m.test(fs.readFileSync('.gitignore', 'utf8')), 'src/appsscript.json ไม่ถูก commit');

const wf = fs.readFileSync('.github/workflows/deploy-appsscript.yml', 'utf8');
assert(/npm install -g @google\/clasp@3/.test(wf), 'ใช้ clasp v3');
assert(/cp appsscript\.json src\/appsscript\.json/.test(wf), 'copy manifest เข้า rootDir ก่อน push');
assert(/clasp push --force/.test(wf), 'push ทับ manifest ได้โดยไม่ถาม');
assert(/clasp deploy --deploymentId "\$DEPLOYMENT_ID"/.test(wf), 'อัปเดต deployment เดิม — URL /exec ไม่เปลี่ยน');
assert(wf.indexOf('Verify webapp access settings') < wf.indexOf('clasp push --force'), 'เช็ก manifest ก่อน push');

console.log('PASS appsscript-deploy-config');
