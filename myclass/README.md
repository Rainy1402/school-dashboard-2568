# MyClass 🍎 — ระบบบันทึกการสอน

เว็บแอปสำหรับครู บันทึกการสอน เช็คชื่อ คะแนน และรายงานผล

## เริ่มต้นใช้งาน (3 ขั้นตอน)

### ขั้น 1 — ตั้งค่า Supabase
1. ไปที่ [supabase.com](https://supabase.com) → สร้าง Project ใหม่ฟรี
2. ไปที่ **SQL Editor** → วางและรันเนื้อหาจากไฟล์ `supabase/schema.sql`
3. ไปที่ **Settings → API** → คัดลอก **Project URL** และ **anon public key**

### ขั้น 2 — รันโปรเจกต์
```bash
npm install
npm run dev
```

### ขั้น 3 — เชื่อมต่อ
เปิดเบราวเซอร์ → กรอก URL และ Key ที่คัดลอกมา → กดเชื่อมต่อ

---

## Deploy บน Vercel (ฟรี)
```bash
npm install -g vercel
vercel --prod
# กรอก VITE_SUPABASE_URL และ VITE_SUPABASE_ANON_KEY ใน Vercel Dashboard
```

## Deploy บน Netlify (ฟรี)
ลาก folder `dist/` (หลัง `npm run build`) ไปที่ [netlify.com/drop](https://netlify.com/drop)

---

## ฟีเจอร์
- ✅ Dashboard สรุปภาพรวม
- ✅ จัดการชั้นเรียน (เพิ่ม/ลบ/กำหนดสี)
- ✅ ฐานข้อมูลนักเรียนกลาง (เพิ่มเอง / วางรายชื่อ / Import Excel)
- ✅ บันทึกการสอน (หัวข้อ จุดประสงค์ กิจกรรม Reflection)
- ✅ กราฟสรุปการเข้าเรียน
- ✅ ตั้งค่าข้อมูลครู/โรงเรียน
