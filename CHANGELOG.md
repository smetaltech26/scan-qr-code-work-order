# 📋 บันทึกการแก้ไขและอัปเดตระบบ (Changelog)

## 🚀 Version 5.2 (อัปเดตล่าสุด: 31 สิงหาคม 2026)

### 🌟 1. เพิ่มเมนูที่ 7: ตั้งค่าระบบ (Settings & Admin Control Panel)
- **🔐 ระบบความปลอดภัย (Password Authentication):**
  - ต้องใส่รหัสผ่านก่อนเข้าหน้า Setting เสมอ (รหัสเริ่มต้น: `1234`)
  - รหัสผ่านบันทึกใน `localStorage` ปลอดภัยและเปลี่ยนได้ตลอดเวลา
- **📦 แท็บ 1: จัดการ Part, สต็อก และราคา (Part & Stock Management):**
  - ค้นหา Part แบบ Real-time Auto-Suggest
  - แก้ไขข้อมูล Part (Name, Model, Drawing, Price ราคาขาย, RM Price ราคาทุน)
  - แก้ไขและบันทึกจำนวนสต็อกแยกทุก Process (`receiving`, `cutting`, `milling`, `cnc`, `grinding`, `finished_goods`, `ng`) พร้อมคำนวณยอด `total` อัตโนมัติ
  - เพิ่ม Part ใหม่ในระบบ พร้อม **ระบบตรวจสอบ ITEM ซ้ำซ้อน (Duplicate Check)** หากซ้ำจะไม่อนุญาตให้เพิ่มและแจ้งเตือนผ่าน Popup ชัดเจน
  - ลบ Part ที่ไม่ต้องการออกจากฐานข้อมูล Supabase
  - ปุ่ม "ล้าง" ข้อมูลการค้นหา
- **📋 แท็บ 2: ประวัติการผลิต (Production Logs Management):**
  - ค้นหาประวัติการสแกนด้วยคำค้นหา (Item, Process, พนักงาน)
  - กรองประวัติการผลิตตามวันที่ (Date Filter)
  - ลบรายการ Log แบบรายตัว (Single Delete)
  - เลือกลบรายการ Log หลายรายการพร้อมกัน (Select All & Bulk Delete)
  - ปุ่ม "ล้าง" ข้อมูลตัวกรอง และปุ่ม "รีเฟรช" ข้อมูลล่าสุด
- **💰 แท็บ 3: ยอดขาย Summary (Daily Summary):**
  - ดึงข้อมูลยอดขายประจำวัน (`daily_income`) และยอดขายสะสม (`accumulate_income`) จากตาราง `daily_summary` ใน Supabase ตามวันที่เลือก
  - บันทึก/อัปเดต (Upsert) ยอดขายของวันที่ต้องการได้อย่างสะดวก
  - แสดงตารางประวัติยอดขายล่าสุด
- **🔑 แท็บ 4: เปลี่ยนรหัสผ่าน Setting (Change Password):**
  - ตรวจสอบรหัสผ่านปัจจุบัน กรอกรหัสใหม่ และยืนยันรหัสใหม่
  - **👁️ เพิ่มปุ่มดวงตา (Password Visibility Toggle):** สามารถกดคลิกเพื่อดู/ซ่อนรหัสผ่านในทุกช่องใส่รหัสผ่านได้อย่างสะดวก

---

### 🎨 2. การปรับปรุง UI/UX และการจัดวาง (Theme & Responsiveness)
- **💎 กู้คืนธีมเดิม 100% (Claymorphism Theme):**
  - หน้าแรก 6 เมนูหลัก: การ์ดทรง Squircle สีพาสเทลนุ่มนวลละมุนตา ความสูงเต็มจอมือถือพอดีเป๊ะตามไฟล์ต้นฉบับ
  - หน้าฟังก์ชันทุกหน้า: คงสีและสไตล์ Claymorphism ดั้งเดิมไว้ครบถ้วน
- **⚙️ ตำแหน่งปุ่ม Setting บนหน้าแรก:**
  - ย้ายปุ่มฟันเฟือง Setting ⚙️ ลงมาอยู่บรรทัดเดียวกับข้อความ `เลือกเมนูเพื่อเริ่มทำงาน [⚙️]`
  - จัดให้อยู่ **กึ่งกลางหน้าจอพอดี (Centered Alignment)** เว้นระยะห่าง 2 เคาะ สวยงามลงตัว
  - หัวเว็บ `Scan QR Code V5.2` โล่ง สบายตา ไม่มีปุ่มใดมาบดบัง
- **📱 ปรับแต่งสำหรับหน้าจอมือถือ (Mobile Compact Optimization):**
  - แถบแท็บ 4 แท็บในหน้า Setting จัดเป็น 4 คอลัมน์กะทัดรัด (Equal 4-Column Grid) ขนาดพอดีหน้าจอ ตัวหนังสือไม่ล้น ไม่เบียด และไม่ตกหล่น
  - ปรับขนาดปุ่ม "ล้าง", "เพิ่ม Part", "กลับหน้าแรก" ให้กะทัดรัด สัดส่วนสวยงาม
- **🖥️ มาตรฐาน Modal เต็มหน้าจอ (Rule 12 Body Portal):**
  - ย้าย Modal ทุกตัว (เพิ่ม Part, ใส่รหัสผ่าน, ยืนยันลบ) ออกมาอยู่นอก Container หลักที่ระดับ `<body>`
  - ฉากหลังมืดโปร่งแสง (Backdrop Blur) แผ่คลุมเต็มหน้าจอ 100vw/100vh ทั้งบนคอมพิวเตอร์และมือถือ
  - ฟอร์ม Modal Scroll ได้อย่างลื่นไหล ปุ่ม Footer ลอยอยู่ด้านล่างชัดเจน ไม่โดนขอบตัดทับ

---

### 🔗 3. รายละเอียดการเชื่อมต่อระบบ (Integration)
- **Production URL:** [https://smetaltech26.github.io/scan-qr-code-work-order/](https://smetaltech26.github.io/scan-qr-code-work-order/)
- **Database:** Supabase Cloud (`orgbrvopqzpfvmtvxmcd.supabase.co`)
  - ตาราง: `inventory_items`, `production_logs`, `daily_summary`
- **GitHub Repository:** `smetaltech26/scan-qr-code-work-order` (Branch: `main`)
