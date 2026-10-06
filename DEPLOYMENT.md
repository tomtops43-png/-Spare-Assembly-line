# Auto-deploy ขึ้น Google Apps Script

โปรเจกต์นี้ตั้ง **GitHub Actions** ให้ deploy โค้ดขึ้น Apps Script อัตโนมัติ
ทุกครั้งที่โค้ดในโฟลเดอร์ `scr/` เข้า `main` (`.github/workflows/deploy-appsscript.yml`)
workflow จะรัน `npm test` + ตรวจสิทธิ์ web app ก่อน แล้วค่อย `clasp push` ขึ้นไป

> ทำตามขั้นตอนข้างล่าง **ครั้งเดียว** เพื่อใส่เครดิต หลังจากนั้นไม่ต้องทำมืออีกเลย —
> ทุก merge จะ `clasp push` ขึ้น Apps Script ให้เอง

---

## ตั้งค่าครั้งเดียว (ทำบนเครื่องคอมพิวเตอร์ของคุณ)

### 1. เปิดใช้ Apps Script API
เข้า https://script.google.com/home/usersettings → เปิด **Google Apps Script API** = **ON**

### 2. ติดตั้ง clasp แล้ว login
```bash
npm install -g @google/clasp@3
clasp login
```
- เลือกบัญชี Google ที่เป็น **เจ้าของโปรเจกต์ Apps Script**
- เสร็จแล้วจะได้ไฟล์เครดิตที่ `~/.clasprc.json`

### 3. Script ID (มีให้แล้ว — ข้ามได้)
repo นี้ commit Script ID ไว้ใน `scr/.clasp.json` แล้ว workflow ใช้ค่านั้นเป็นค่าตั้งต้น
ใส่ secret `SCRIPT_ID` เฉพาะตอนอยากชี้ไปโปรเจกต์ Apps Script ตัวอื่น
(หาได้ที่ Apps Script → ⚙️ **Project Settings** → **Script ID**)

### 4. (ถ้าต้องการให้ URL `/exec` เดิมไม่เปลี่ยน) หา Deployment ID
ในโปรเจกต์ Apps Script → **Deploy → Manage deployments** → คัดลอก **Deployment ID** ของ web app

### 5. ใส่ Secrets ใน GitHub
ไปที่ repo → **Settings → Secrets and variables → Actions → New repository secret**
แล้วเพิ่ม:

| ชื่อ Secret | ค่า | จำเป็น |
|-------------|-----|--------|
| `CLASPRC_JSON` | เนื้อหาทั้งหมดของไฟล์ `~/.clasprc.json` | ✅ ใช่ |
| `SCRIPT_ID` | Script ID จากข้อ 3 | ⬜ ไม่ต้อง — ใช้ค่าใน `scr/.clasp.json` |
| `DEPLOYMENT_ID` | Deployment ID จากข้อ 4 | ⬜ ถ้าอยากให้ URL เดิมคงที่ |

วิธีก๊อปเนื้อหา `~/.clasprc.json`:
```bash
# mac / linux
cat ~/.clasprc.json
# windows (cmd)
type %USERPROFILE%\.clasprc.json
```
ก๊อปทั้งก้อน (รวม `{ ... }`) ไปวางเป็นค่าของ secret `CLASPRC_JSON`

---

## เสร็จแล้ว
- ทุกครั้งที่มีโค้ดใน `scr/` เข้า `main` → ดู progress ได้ที่แท็บ **Actions**
- อยากสั่ง deploy เองตอนไหนก็ได้: **Actions → Deploy to Apps Script → Run workflow**

## สิทธิ์เข้าถึง web app — "ทุกคนที่มีลิงก์"

แอปนี้ต้องเปิดให้ **ทุกคนที่มีลิงก์** เข้าได้ เพราะเครื่องหน้างานไม่ได้ล็อกอิน Google
และหน้าเว็บเรียก `/exec` แบบ `credentials: 'omit'` ถ้าตั้งเป็นอย่างอื่น ทุก request
จะโดนเด้งไปหน้า login ของ Google ซึ่งเบราว์เซอร์รายงานกลับมาเป็น CORS error —
อาการที่เห็นคือแอปพัง ไม่ใช่แอปถูกล็อก

ค่านี้อยู่ใน `scr/appsscript.json` และถูก push ขึ้นไปทุกครั้งที่ deploy:

```json
"webapp": {
  "access": "ANYONE_ANONYMOUS",
  "executeAs": "USER_DEPLOYING"
}
```

| ค่า | ความหมายในหน้า Apps Script | ใช้ได้ไหม |
|-----|---------------------------|-----------|
| `ANYONE_ANONYMOUS` | **Anyone** — ไม่ต้องล็อกอิน = ทุกคนที่มีลิงก์ | ✅ ค่าที่ต้องใช้ |
| `ANYONE` | Anyone with a Google account — ต้องล็อกอินก่อน | ❌ แอปจะพัง |
| `DOMAIN` / `MYSELF` | เฉพาะในองค์กร / เฉพาะเจ้าของ | ❌ แอปจะพัง |

`executeAs: USER_DEPLOYING` ต้องคู่กันเสมอ — สคริปต์รันด้วยสิทธิ์บัญชีเจ้าของ
คนที่เปิดแอปแบบไม่ล็อกอินจึงเข้าถึง Google Sheet ได้ผ่านสคริปต์

workflow มีขั้นตอน **Verify web app stays public** ตรวจสองค่านี้ก่อน push (และมีเทสต์
`tests/appsscript-webapp-public.test.js` จับตั้งแต่ตอนเปิด PR) ถ้าใครแก้เป็น
ค่าอื่น deploy จะล้มพร้อมข้อความบอกสาเหตุ แทนที่จะเงียบ ๆ ปล่อยแอปที่เข้าไม่ได้ขึ้น production

### ถ้า deployment เดิมยังไม่เปิดเป็น "ทุกคนที่มีลิงก์"
`clasp deploy --deploymentId` อัปเดต **โค้ด** ของ deployment เดิม แต่ไม่ย้อนกลับไปแก้สิทธิ์
ที่ตั้งไว้ตอนสร้าง deployment นั้นครั้งแรก ถ้าเปิด `/exec` แล้วยังเจอหน้า login ของ Google
ต้องแก้ **ครั้งเดียว** ด้วยมือ:

1. เปิดโปรเจกต์ Apps Script → **Deploy → Manage deployments**
2. กดดินสอ ✏️ ที่ deployment ของ web app
3. **Who has access** → เลือก **Anyone**
4. **Deploy**

URL `/exec` ไม่เปลี่ยน และหลังจากนี้ทุก merge จะอัปเดตโค้ดให้เองตามเดิม

---

## หมายเหตุ
- ถ้าไม่ได้ใส่ `CLASPRC_JSON` workflow จะล้มที่ขั้นตอน **Check secrets** พร้อมบอกให้มาอ่านไฟล์นี้
- ถ้า **ไม่ใส่** `DEPLOYMENT_ID`: โค้ดจะถูก push ขึ้น HEAD เท่านั้น ถ้า URL `/exec`
  ของคุณตรึงไว้ที่เวอร์ชันใดเวอร์ชันหนึ่ง ต้องใส่ `DEPLOYMENT_ID` ด้วย URL ถึงจะอัปเดตตาม
- ถ้า token หมดอายุ/ถูกเพิกถอน ให้รัน `clasp login` ใหม่ แล้วอัปเดต secret `CLASPRC_JSON`
- workflow ตรึง `clasp@3` ไว้ (เอาแค่ major กันพังจากการอัปเดตข้ามเวอร์ชันใหญ่)
  ตอน `clasp login` บนเครื่องตัวเองต้องใช้ major เดียวกันด้วย เพราะ clasp คนละ major
  เขียน `~/.clasprc.json` คนละรูปแบบ — ถ้าใช้คนละเวอร์ชัน CI จะขึ้น
  `Error retrieving access token`
