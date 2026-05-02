# ระบบคะแนนพฤติกรรมนักเรียน
## โรงเรียนด่านช้างวิทยา
### พัฒนาโดย นายณัฐพล กลุ่มกลัด

---

## ⚡ ระบบนี้ใช้ Vite (เร็วกว่า Create React App มาก, warning น้อยมาก)

---

## ขั้นตอนที่ 1 — ติดตั้ง Node.js

1. ไปที่ **https://nodejs.org** → ดาวน์โหลด **LTS**
2. ติดตั้งตามปกติ กด Next ไปเรื่อยๆ
3. ตรวจสอบ: เปิด **Command Prompt** → พิมพ์ `node -v`
   - ✅ ขึ้นเลขเวอร์ชัน = สำเร็จ

---

## ขั้นตอนที่ 2 — ติดตั้ง Dependencies

เปิด **Command Prompt** แล้วพิมพ์:

```
cd C:\Users\ชื่อคุณ\school-behavior
npm install
```

✅ **Vite จะมี warning น้อยมาก** ต่างจาก Create React App

---

## ขั้นตอนที่ 3 — ตั้งค่า Firebase

### 3.1 สร้าง Firebase Project

1. ไปที่ **https://console.firebase.google.com**
2. Sign in ด้วย Gmail → กด **"Create a project"**
3. ตั้งชื่อ: `danchan-behavior` → Create project

### 3.2 เปิด Firestore Database

1. เมนูซ้าย → **Build** → **Firestore Database**
2. กด **"Create database"** → เลือก **"Start in test mode"**
3. Region: **`asia-southeast1`** → Enable

### 3.3 ก็อป Firebase Config

1. ⚙️ → **"Project settings"** → **"Your apps"** → กด **`</>`**
2. ตั้งชื่อ App → Register app → ก็อป firebaseConfig

### 3.4 วาง Config ในไฟล์

เปิดไฟล์ `src/firebase.js` ด้วย Notepad  
หาบรรทัด `✏️ แก้ค่าตรงนี้` แล้วแทนที่ค่า firebaseConfig  
บันทึกไฟล์ (Ctrl+S)

---

## ขั้นตอนที่ 4 — แก้รหัสผ่าน (ถ้าต้องการ)

เปิด `src/App.jsx` หาบรรทัด `CREDENTIALS`:

```js
const CREDENTIALS = {
  admin:   { password: 'admin2567',   ... },
  teacher: { password: 'teacher2567', ... },
}
```

เปลี่ยนรหัสผ่านตามต้องการ บันทึกไฟล์

---

## ขั้นตอนที่ 5 — ทดสอบบนเครื่อง

```
npm run dev
```

เปิดเบราว์เซอร์ไปที่ **http://localhost:5173**

ทดสอบ:
- Login admin / teacher
- นำเข้า Excel
- บันทึกคะแนน

---

## ขั้นตอนที่ 6 — Deploy ขึ้น Vercel (ฟรี)

### สมัคร GitHub + อัปโหลด

```
git init
git add .
git commit -m "first deploy"
git branch -M main
git remote add origin https://github.com/nattapon29-29/school-behavior.git
git push -u origin main
```

### Deploy ด้วย Vercel

1. **https://vercel.com** → Sign up with GitHub
2. **"Add New Project"** → เลือก repo
3. กด **"Deploy"** → รอ 2 นาที
4. ✅ ได้ลิงก์: `https://school-behavior-xxx.vercel.app`

---

## ขั้นตอนที่ 7 — ใช้งานจริง

### นำเข้ารายชื่อนักเรียน 700 คน
1. Login Admin → **"นำเข้า Excel"** → ดาวน์โหลด Template
2. กรอกรายชื่อใน Excel → อัปโหลด → **"แทนที่ทั้งหมด"**

### ตั้งค่าเทอม
Admin → แท็บ **"จัดการเทอม"** → สลับไปเทอมที่ถูกต้อง

### แจกให้ครู 40 คน
ส่ง LINE:
```
🏫 ระบบคะแนนพฤติกรรม โรงเรียนด่านช้างวิทยา
🔗 https://school-behavior-xxx.vercel.app
👩‍🏫 เลือก "ครู" → รหัสผ่าน: teacher2567
📱 Add to Home Screen เพื่อใช้เหมือนแอป
```

---

## 🗓️ ปฏิทินการใช้งาน

| ช่วงเวลา | สิ่งที่ต้องทำ |
|:--------|:------------|
| ต้นเทอม 1 | นำเข้ารายชื่อนักเรียน |
| ระหว่างเทอม | ครูบันทึกคะแนนตามปกติ |
| สิ้นเทอม 1 | Export รายงาน |
| ต้นเทอม 2 | กด "เลื่อนขึ้นเทอมใหม่" |
| สิ้นปีการศึกษา | Export → เคลียข้อมูล |

---

## 🔑 รหัสผ่านเริ่มต้น

| บทบาท | รหัสผ่าน |
|:-----:|:-------:|
| แอดมิน | admin2567 |
| ครู (ทุกคน) | teacher2567 |

แก้ได้ที่ `src/App.jsx` บรรทัด `CREDENTIALS`

---

## ❓ แก้ปัญหาที่พบบ่อย

| ปัญหา | วิธีแก้ |
|:------|:-------|
| `npm install` มี warning | ✅ ปกติมาก กด Enter ผ่านได้เลย |
| Firebase error | ตรวจสอบ firebaseConfig ใน `src/firebase.js` |
| ข้อมูลไม่บันทึก | ตรวจสอบ Firestore อยู่ใน Test mode |
| หน้าเว็บค้าง | กด F12 → Console → ดู error แดง |
| Vercel build fail | ตรวจว่าทุกไฟล์ push ขึ้น GitHub แล้ว |

---

*พัฒนาด้วย React 18 + Vite 5 + Firebase 10*
*ฟรี 100% ไม่มีค่าใช้จ่าย*
