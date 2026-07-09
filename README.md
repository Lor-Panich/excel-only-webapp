# vRich Stock Excel Update Static Web App

เว็บแอพนี้เป็น Static Web App สำหรับ GitHub Pages และทำงานแบบ browser-only ทั้งหมด

- ไม่ใช้ API
- ไม่ใช้ Console scraper
- ไม่ดึงเว็บเพื่ออ่านข้อมูลบริษัท
- ไม่มี backend
- ไม่อัปโหลดไฟล์ไป server
- ไม่ import เข้า vRich เอง
- ประมวลผลไฟล์ Excel ใน browser ของผู้ใช้เท่านั้น

## วิธีใช้งานเว็บ

1. เปิด `index.html`
2. เลือกไฟล์ 3 ไฟล์:
   - `stock_vrich619.xlsx`
   - `Item jsterp.xlsx`
   - `sold_today.xlsx`
3. กด `ตรวจสอบไฟล์`
4. ตรวจ preflight ว่า sheet, columns, row count ถูกต้อง
5. กด `ประมวลผล`
6. ถ้ามี missing เช่น `FN913` ให้ติ๊ก exclude เฉพาะเมื่อยืนยันแล้วว่าจะไม่อัปเดตรหัสนั้น
7. ดาวน์โหลดไฟล์ผลลัพธ์และ reports

Default config:

```text
SOLD_TODAY_CODE_COLUMN = รหัสสินค้า
VRICH_MATCH_COLUMN = รหัสขาย
JST_MATCH_COLUMN = รหัสรูปแบบ
VRICH_QTY_COLUMN = จำนวน
JST_QTY_COLUMN = จำนวน
```

## วิธีรัน local

เปิดไฟล์นี้ใน browser ได้โดยตรง:

```text
index.html
```

หรือรัน static server:

```bash
python -m http.server 8000
```

แล้วเปิด:

```text
http://localhost:8000
```

## วิธีเอาขึ้น GitHub Pages

1. สร้าง repository สำหรับเว็บแอพ
2. commit เฉพาะ source code:
   - `index.html`
   - `app.js`
   - `style.css`
   - `README.md`
   - `.gitignore`
   - `vendor/xlsx.full.min.js`
3. เข้า GitHub repository settings
4. เปิด Pages
5. เลือก deploy จาก branch เช่น `main` และ folder root

## คำเตือนเรื่องไฟล์ Excel จริง

ห้าม commit ไฟล์ Excel จริงของบริษัทขึ้น GitHub โดยเด็ดขาด

`.gitignore` ในโปรเจกต์นี้ block ไฟล์ Excel และ output reports เช่น:

- `*.xlsx`
- `*.xls`
- `vrich_import_update_qty.xlsx`
- `summary_report.xlsx`
- `report_*.xlsx`
- `excluded_codes_report.xlsx`

## Library ภายนอก

โปรเจกต์นี้ vendor `SheetJS xlsx.full.min.js` ไว้ใน `vendor/` แล้ว จึงไม่ต้องโหลด CDN ตอนใช้งานเว็บ

## ขั้นตอนตรวจก่อน import เข้า vRich

1. เปิด `summary_report.xlsx`
2. ตรวจ status:
   - `PASS`: ไม่มี missing และไม่มี duplicate ที่ชนกับ sold_today
   - `PASS_WITH_EXCLUSION`: ไม่มี missing หลังผู้ใช้ยืนยัน exclude แล้ว
   - `FAIL`: ยังไม่ควร import
   - `FAIL_DUPLICATE`: ยังไม่ควร import เพราะมีรหัสซ้ำชนกับ sold_today
3. เปิด `report_missing_in_vrich.xlsx`
4. เปิด `report_missing_in_jst.xlsx`
5. ถ้ามี `excluded_codes_report.xlsx` ให้ยืนยันอีกครั้งว่ารหัสเหล่านั้นไม่ต้องอัปเดต
6. เปิด `vrich_import_update_qty.xlsx` และตรวจจำนวนแถวกับคอลัมน์
7. import เข้า vRich ด้วยตัวเองเท่านั้น เมื่อมั่นใจแล้ว
