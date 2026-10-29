# Template: Auto-deploy เว็บ (GitHub Pages) + backend (Google Apps Script)

ก๊อปทั้งโฟลเดอร์นี้ไปวางที่ **root ของ repo ใหม่** แล้วทำตาม "ตั้งค่าครั้งเดียว" ข้างล่าง
หลังจากนั้นแค่ push เข้า `main` ระบบจะ deploy ให้เองทั้งสองฝั่ง

```
repo-ใหม่/
├── .github/workflows/
│   ├── deploy-pages.yml        ← push เข้า main → ขึ้น GitHub Pages
│   └── deploy-appsscript.yml   ← แก้ไฟล์ใน gas/ แล้วเข้า main → clasp push ขึ้น Apps Script
├── gas/
│   ├── .clasp.json             ← ใส่ Script ID
│   ├── appsscript.json         ← manifest (สิทธิ์ web app = Anyone)
│   └── Code.gs                 ← ตัวอย่าง doGet/doPost — แทนที่ด้วยโค้ดของคุณ
└── index.html                  ← หน้าเว็บของคุณ (template ไม่ได้ให้มา)
```

ก๊อปจาก repo นี้ไป repo ใหม่ (สมมติว่า clone ทั้งสองไว้ข้างกัน):

```bash
cp -r -Spare-Assembly-line/templates/auto-deploy/. repo-ใหม่/
```

> ใช้แค่ฝั่งเดียวก็ได้ — มีแค่หน้าเว็บก็ลบ `deploy-appsscript.yml` + `gas/` ทิ้ง
> มีแค่ GAS ก็ลบ `deploy-pages.yml` ทิ้ง

---

## ส่วนที่ 1: GitHub Pages (ทำครั้งเดียว)

1. repo ใหม่ → **Settings → Pages → Build and deployment → Source** = **GitHub Actions**
2. push เข้า `main` → ดูความคืบหน้าที่แท็บ **Actions**
3. เว็บจะอยู่ที่ `https://<user>.github.io/<repo>/`

workflow คัดลอกทุกไฟล์ใน repo ขึ้นเว็บ **ยกเว้น** `.github/`, `gas/`, `tests/`, `node_modules/`,
`package*.json`, `*.md` — ถ้าเว็บอยู่ในโฟลเดอร์เดียว (เช่น `public/`) แก้ขั้น
`Prepare static site artifact` ใน `deploy-pages.yml` ตามคอมเมนต์

ถ้ามี `package.json` ที่มีสคริปต์ `test` จะรัน `npm test` ก่อน deploy (เทสต์แดง = ไม่ deploy)
ไม่มีก็ข้ามไป

---

## ส่วนที่ 2: Google Apps Script (ทำครั้งเดียว บนคอมของคุณ)

### 1. เปิด Apps Script API
https://script.google.com/home/usersettings → **Google Apps Script API** = **ON**

### 2. สร้างโปรเจกต์ GAS แล้วเอา Script ID
- สร้างโปรเจกต์ใหม่ที่ https://script.google.com (หรือเปิด Google Sheet → ส่วนขยาย → Apps Script)
- ⚙️ **Project Settings** → คัดลอก **Script ID**
- ใส่ใน `gas/.clasp.json` แทน `ใส่_SCRIPT_ID_ตรงนี้`

### 3. Deploy เป็น web app ครั้งแรก (มือ ครั้งเดียว)
ในหน้า Apps Script → **Deploy → New deployment** → ประเภท **Web app**
- **Execute as**: Me
- **Who has access**: **Anyone**
- กด Deploy → คัดลอก **Web app URL** (`.../exec`) ไปใส่ในหน้าเว็บ
- **Deploy → Manage deployments** → คัดลอก **Deployment ID** (ใช้ในข้อ 5)

### 4. ติดตั้ง clasp แล้ว login

```bash
npm install -g @google/clasp@3
```

```bash
clasp login
```

เลือกบัญชีที่เป็นเจ้าของโปรเจกต์ GAS → ได้ไฟล์ `~/.clasprc.json`

> ต้องใช้ clasp **major 3** ให้ตรงกับ workflow — คนละ major เขียนไฟล์เครดิตคนละรูปแบบ
> CI จะขึ้น `Error retrieving access token`

### 5. ใส่ Secrets ใน GitHub
repo ใหม่ → **Settings → Secrets and variables → Actions → New repository secret**

| ชื่อ Secret | ค่า | จำเป็น |
|-------------|-----|--------|
| `CLASPRC_JSON` | เนื้อหาทั้งไฟล์ `~/.clasprc.json` (รวม `{ ... }`) | ✅ |
| `DEPLOYMENT_ID` | Deployment ID จากข้อ 3 | ✅ ถ้าอยากให้ URL `/exec` ได้โค้ดใหม่ทุกครั้ง |
| `SCRIPT_ID` | Script ID | ⬜ ไม่ต้อง ถ้าใส่ใน `gas/.clasp.json` แล้ว |

ดูเนื้อหาไฟล์เครดิต:

```bash
cat ~/.clasprc.json
```

(Windows cmd: `type %USERPROFILE%\.clasprc.json`)

> ⚠️ `CLASPRC_JSON` คือกุญแจเข้าบัญชี Google ของคุณ — ใส่ใน Secrets เท่านั้น **ห้าม commit ลง repo**

### เสร็จแล้ว
แก้ไฟล์ใน `gas/` → merge เข้า `main` → workflow จะ `clasp push` + `clasp deploy` ให้เอง
URL `/exec` ไม่เปลี่ยน สั่งเองได้ที่ **Actions → Deploy to Apps Script → Run workflow**

---

## เรียก GAS จากหน้าเว็บ

```js
const API_URL = 'https://script.google.com/macros/s/xxxxxxxx/exec';

// GET
const res = await fetch(API_URL + '?action=ping');
console.log(await res.json());

// POST — ใช้ text/plain เพื่อไม่ให้เบราว์เซอร์ยิง preflight (GAS ไม่ตอบ OPTIONS → CORS error)
const res2 = await fetch(API_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain;charset=utf-8' },
  body: JSON.stringify({ action: 'ping' })
});
console.log(await res2.json());
```

---

## แก้ปัญหา

| อาการ | สาเหตุ / วิธีแก้ |
|-------|------------------|
| ล้มที่ **Check secrets** | ยังไม่ได้ใส่ `CLASPRC_JSON` |
| `Error retrieving access token` | token หมดอายุ หรือ login ด้วย clasp คนละ major → รัน `clasp login` (v3) ใหม่ แล้วอัปเดต secret |
| `User has not enabled the Apps Script API` | ยังไม่ได้ทำข้อ 1 |
| ล้มที่ **Verify web app stays public** | `gas/appsscript.json` ไม่ได้ตั้ง `ANYONE_ANONYMOUS` / `USER_DEPLOYING` |
| หน้าเว็บเรียก `/exec` แล้วเจอ CORS error | web app ไม่ได้เปิด **Anyone** → Manage deployments → ✏️ → Who has access = **Anyone** → Deploy (`clasp deploy` ไม่แก้สิทธิ์ของ deployment ที่สร้างไว้แล้ว) |
| deploy ผ่านแต่ `/exec` ยังเป็นโค้ดเก่า | ไม่ได้ใส่ `DEPLOYMENT_ID` |
| เพิ่ม scope ใหม่ใน `appsscript.json` แล้ว GAS error เรื่องสิทธิ์ | เปิดโปรเจกต์ใน Apps Script แล้วกด Run ฟังก์ชันใดก็ได้หนึ่งครั้งเพื่อกดอนุญาต scope ใหม่ |
| Pages ล้มที่ **Deploy to GitHub Pages** | ยังไม่ได้ตั้ง Settings → Pages → Source = GitHub Actions |
