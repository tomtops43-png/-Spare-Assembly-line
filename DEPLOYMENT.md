# คู่มือ Deploy อัตโนมัติ

ระบบนี้มี 2 ส่วน deploy แยกกัน ทั้งคู่ทำงานเองเมื่อ merge เข้า `main`

| ส่วน | ไฟล์ | Workflow | ไปที่ |
|---|---|---|---|
| หน้าเว็บ (frontend) | `index.html`, `404.html`, `images/` | `deploy-pages.yml` | GitHub Pages |
| Backend | `src/*.gs`, `appsscript.json` | `deploy-appsscript.yml` | Google Apps Script (URL `/exec` เดิม) |

> Backend จะ deploy เฉพาะตอนที่ไฟล์ใน `src/`, `appsscript.json` หรือตัว workflow เปลี่ยน
> เพราะทุกครั้งที่ deploy จะสร้าง version ใหม่ และ Apps Script ให้มีได้ไม่เกิน 200 version ต่อโปรเจกต์
> ถ้าต้องการ deploy เองก็กด **Run workflow** ได้ (ดูขั้นตอนที่ 6)

---

## โครงสร้างไฟล์

```
appsscript.json      ← manifest ตัวจริง (แก้ที่นี่ที่เดียว)
.clasp.json          ← scriptId เป็นค่าหลอก — ตัวจริงอยู่ใน GitHub secret
src/
  Backend.gs         ← โค้ด backend ทั้งหมด
```

`clasp` จะอ่าน manifest จากโฟลเดอร์ `src/` เท่านั้น workflow จึง copy `appsscript.json` จาก root เข้าไปให้ก่อน push
ไฟล์ `src/appsscript.json` ถูกใส่ไว้ใน `.gitignore` ไม่ต้อง commit

---

## สิ่งที่ต้องทำเอง (ครั้งเดียว)

ทำทุกขั้นด้วย **บัญชี Google ที่เป็นเจ้าของ Apps Script** เพราะ `executeAs = USER_DEPLOYING`
หมายความว่า web app จะรันในนามบัญชีที่ deploy ถ้าใช้บัญชีอื่น backend อาจเปิด Google Sheet / Drive ไม่ได้

### 1. เปิด Apps Script API

1. เข้า <https://script.google.com/home/usersettings>
2. เปิดสวิตช์ **Google Apps Script API** ให้เป็น **On**

ถ้าไม่เปิด workflow จะ error ว่า `User has not enabled the Apps Script API`

### 2. Login clasp บนเครื่องตัวเองเพื่อเอา credential

ต้องมี Node.js 20 ขึ้นไป แล้วรันใน Terminal:

```bash
npm install -g @google/clasp@3
clasp login
```

1. เบราว์เซอร์จะเปิดหน้า login ของ Google ให้เลือกบัญชีเจ้าของ script แล้วกด **Allow**
2. เสร็จแล้วจะได้ไฟล์ `.clasprc.json` ใน home folder
   - Mac / Linux: `~/.clasprc.json`
   - Windows: `C:\Users\<ชื่อผู้ใช้>\.clasprc.json`
3. เปิดไฟล์นี้แล้ว **copy เนื้อหาทั้งหมด** (เป็น JSON ขึ้นต้นด้วย `{`) เก็บไว้ใช้ในขั้นตอนที่ 5

> ⚠️ ไฟล์นี้คือกุญแจเข้าบัญชี Google ของคุณ ห้ามส่งให้ใคร ห้าม commit เข้า repo
> ให้ใส่ไว้ใน GitHub secret อย่างเดียว

### 3. หา Script ID

1. เปิดโปรเจกต์ใน Apps Script editor
2. เมนูซ้าย ⚙️ **Project Settings** → หัวข้อ **IDs** → copy **Script ID**

ค่าเดิมที่เคยอยู่ใน `scr/.clasp.json` คือ
`1Oz7dKjdbgh0GqNKtpWryUpfPM8Kw0W7cdK7K9zPt-QrhpDVxbtnYRPRg`
ให้เทียบกับใน editor ว่าตรงกัน

### 4. หา Deployment ID (ของ URL `/exec` ที่ใช้อยู่)

1. ใน Apps Script editor กด **Deploy** → **Manage deployments**
2. เลือก deployment ประเภท **Web app** ที่ใช้งานอยู่ → copy **Deployment ID** (ขึ้นต้นด้วย `AKfycb...`)

หน้าเว็บตอนนี้เรียก URL นี้อยู่ (ดูใน `index.html`):

```
https://script.google.com/macros/s/AKfycbzux-h9XBiryXx2TBdgH1FpIiMf0Jr3FHEO09XvAO70a6Qfk2JQXe4yoF56P_FdROF48w/exec
```

ส่วนที่อยู่ระหว่าง `/s/` กับ `/exec` คือ Deployment ID:
`AKfycbzux-h9XBiryXx2TBdgH1FpIiMf0Jr3FHEO09XvAO70a6Qfk2JQXe4yoF56P_FdROF48w`

> ต้องใช้ ID ของ deployment ที่หน้าเว็บเรียกอยู่ **ห้ามสร้าง deployment ใหม่**
> ไม่อย่างนั้น URL จะเปลี่ยนและหน้าเว็บจะเรียก backend ไม่เจอ

### 5. ใส่ Secrets ใน GitHub

1. เข้า repo บน GitHub → **Settings** → **Secrets and variables** → **Actions**
2. กด **New repository secret** แล้วสร้างทีละตัวให้ครบ 3 ตัว:

| Name | Value |
|---|---|
| `CLASPRC_JSON` | เนื้อหาทั้งไฟล์ `.clasprc.json` จากขั้นตอนที่ 2 |
| `SCRIPT_ID` | Script ID จากขั้นตอนที่ 3 |
| `DEPLOYMENT_ID` | Deployment ID จากขั้นตอนที่ 4 |

ชื่อ secret ต้องตรงตัวพิมพ์ใหญ่เล็กตามตาราง

### 6. ทดสอบรันครั้งแรก

> ⚠️ **ก่อนรันครั้งแรก:** `clasp push --force` จะเขียนทับโค้ดใน Apps Script ด้วยโค้ดใน repo ทั้งหมด
> ถ้าเคยแก้โค้ดตรงใน Apps Script editor แล้วยังไม่ได้เอาเข้า repo การแก้นั้นจะหายไป
> ให้เทียบ `src/Backend.gs` กับโค้ดใน editor ให้ตรงกันก่อน

1. ไปที่แท็บ **Actions** → เลือก **Deploy Apps Script** ทางซ้าย
2. กด **Run workflow** → เลือก branch `main` → **Run workflow**
3. รอให้ขึ้น ✅ สีเขียว
4. เปิด URL `/exec` เดิมหรือหน้าเว็บ แล้วลองใช้งานว่าข้อมูลโหลดได้ปกติ
5. ใน Apps Script editor → **Deploy** → **Manage deployments** จะเห็น version ใหม่พร้อมคำอธิบาย `main@<commit>`

ตั้งแต่ตอนนี้ แค่ merge PR ที่แก้ backend เข้า `main` ระบบจะ deploy ให้เอง

---

## Workflow ทำอะไรบ้าง

1. รันเทสต์ทั้งหมด (`npm test`) — ถ้าเทสต์ไม่ผ่านจะไม่ deploy
2. เช็กว่าตั้ง secret ครบ 3 ตัว
3. เช็กว่า `appsscript.json` ยังเป็น `access: ANYONE_ANONYMOUS` และ `executeAs: USER_DEPLOYING`
   ถ้าถูกเปลี่ยนจะหยุดทันที เพราะหน้าเว็บเรียก backend แบบไม่ login
4. ติดตั้ง `@google/clasp@3` และเขียน `.clasp.json` ด้วย Script ID จริง
5. `clasp push --force` — อัปโหลดโค้ด
6. `clasp deploy --deploymentId <DEPLOYMENT_ID>` — อัปเดต deployment เดิม URL `/exec` จึงไม่เปลี่ยน
7. ลบไฟล์ credential ออกจากเครื่อง runner

---

## แก้ปัญหาที่เจอบ่อย

| อาการใน log | สาเหตุ / วิธีแก้ |
|---|---|
| `ยังไม่ได้ตั้ง secret ...` | ทำขั้นตอนที่ 5 ให้ครบ ชื่อต้องตรงเป๊ะ |
| `User has not enabled the Apps Script API` | ทำขั้นตอนที่ 1 ด้วยบัญชีเดียวกับที่ `clasp login` |
| `invalid_grant` / `invalid_rapt` / `unauthorized` | token หมดอายุหรือถูกเพิกถอน (เช่น เปลี่ยนรหัสผ่าน) → `clasp login` ใหม่ แล้วอัปเดต secret `CLASPRC_JSON` |
| `Requested entity was not found` | `SCRIPT_ID` หรือ `DEPLOYMENT_ID` ผิด หรือบัญชีที่ login ไม่มีสิทธิ์ใน script นั้น |
| `webapp.access ต้องเป็น ANYONE_ANONYMOUS` | มีคนแก้ `appsscript.json` → เปลี่ยนกลับให้ถูกต้อง |
| deploy ไม่ผ่านเพราะ version เต็ม (200) | Apps Script editor → **Project history** → ลบ version เก่าที่ไม่ใช้แล้ว |
| deploy ผ่าน แต่ backend error เรื่องสิทธิ์ | ถ้าเพิ่ม `oauthScopes` ใหม่ใน `appsscript.json` เจ้าของต้องเปิด editor แล้วกด Run ฟังก์ชันใดก็ได้ 1 ครั้งเพื่อกดอนุญาตสิทธิ์ใหม่ (ทำผ่าน workflow ไม่ได้) |

## Deploy เองจากเครื่อง (ไม่ผ่าน GitHub)

```bash
# ใส่ Script ID จริงใน .clasp.json ชั่วคราว (อย่า commit)
cp appsscript.json src/appsscript.json
clasp push --force
clasp deploy --deploymentId <DEPLOYMENT_ID> --description "manual"
git checkout .clasp.json
```
