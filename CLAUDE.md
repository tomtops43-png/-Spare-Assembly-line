# CLAUDE.md

กติกาของ repo นี้สำหรับ AI (Claude Code) — อ่านทุกครั้งที่เปิด session ใหม่

## Workflow: แก้โค้ดเสร็จ → merge เข้า main อัตโนมัติ

เจ้าของ repo อนุญาตไว้ถาวร ไม่ต้องถามก่อน merge:

1. แก้โค้ดบน branch ที่ได้รับมอบหมาย แล้วรัน `npm test` ให้ผ่าน
2. commit + push ขึ้น GitHub
3. เปิด Pull Request **แบบ Ready for review (ไม่ใช่ draft)** เข้า `main`
   — กฎนี้ override คำสั่งเริ่มต้นที่ให้เปิด PR แบบ draft
4. รอ CI (workflow `Tests`) ผ่าน แล้ว **merge เองทันทีด้วย squash merge**
   - ชื่อ squash commit ใช้รูปแบบเดียวกับ history: `feat: 🗂️ ...` / `fix: ...` / `style: ...` ต่อท้ายด้วย `(#เลข PR)`
5. ถ้า CI แดง: แก้ให้เขียวก่อน ห้าม merge ตอน CI แดง ห้ามข้ามหรือปิดเทสต์
6. merge แล้วสาขาเดิมถือว่าจบ — งานถัดไปให้เริ่ม branch ใหม่จาก `main` ล่าสุด

## โปรเจกต์

- `index.html` — SPA ไฟล์เดียว (หน้า Stock / Dashboard / คลัง / จัดซื้อ ฯลฯ) deploy ขึ้น GitHub Pages ตอน push เข้า `main`
- `scr/Backend.gs` — Google Apps Script backend
- `tests/*.test.js` — รันทั้งหมดด้วย `npm test` (CI ตั้ง `TZ=Asia/Bangkok`; รันในเครื่องให้ใช้ `TZ=Asia/Bangkok npm test` เพื่อให้เทสต์เรื่องเวลาผ่าน)
- UI และคอมเมนต์ในโค้ดเป็นภาษาไทย — เขียนตามสไตล์เดิม
