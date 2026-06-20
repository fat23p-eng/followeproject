ลองทำตามขั้นตอนนี้ได้เลย:
=========
drive.google.com → + New → Google Sheets
ใน Sheet → เมนู Extensions → Apps Script
วางโค้ด Code.gs และ Index.html
รัน setup()

ปีงบประมาณหน้า จะแก้ไขตรงไหนครับ
=========
13 พ.ค.
แก้ไขแค่ 2 จุด ใน Index.html ครับ
จุดที่ 1 — Navbar title (บรรทัดประมาณ 9)
html
<title>ระบบติดตามผลการปฏิบัติงาน 2570</title>

จุดที่ 2 — Navbar subtitle
html
<div class="navbar-sub">ปีงบประมาณ 2570</div>
และใน Code.gs อีก 1 จุด

จุดที่ 3 — doGet title
javascript
.setTitle('ระบบติดตามผลการปฏิบัติงาน 2570')
💡 แนะนำ — ปีหน้าควรสร้าง Google Sheet ใหม่อีกไฟล์หนึ่งแยกออกไปเลย เพื่อเก็บข้อมูลปี 2569 ไว้เป็นประวัติ แล้ว Deploy ระบบใหม่สำหรับปี 2570 ครับ






