# B-L-CHECK

โปรแกรมตรวจ B/L ก่อนออก B/L จริง — เทียบ **B/L ของลูกค้า** (`CLIENT … .pdf`) กับ **DRAFT ของ HAL**
(`DRAFT HAL … .pdf`) ครบ 27 หัวข้อ รวม **RIDER / AS PER ATTACHED** แล้วออกรายงาน
**`BL CHECK <เลข B/L>.xlsx`** แบบใส่สี

- B/L NO., SHIPPER, CONSIGNEE, NOTIFY PARTY, PRE CARRIAGE BY, PLACE OF RECEIPT
- OCEAN VESSEL VOY NO., PORT OF LOADING, PORT OF DISCHARGE, PLACE OF DELIVERY, FINAL DESTINATION
- CONTAINER NO. / SEAL NO., NO. OF CONTAINERS OR PKGS, MARKS AND NUMBERS, DESCRIPTION OF GOODS
- GROSS WEIGHT, MEASUREMENT, TOTAL (IN WORDS), FREIGHT AND CHARGES, EX RATE, PREPAID AT, PAYABLE AT
- PLACE AND DATE OF ISSUE, TOTAL PREPAID IN LOCAL CURRENCY, NO. OF ORIGINAL B(S)/L, DATE / LADEN ON BOARD, CY/CY
- RIDER / ATTACHED SHEET: เทียบทีละบรรทัด ทีละตู้ และยอดรวม KGS · ถ้าฝั่งหนึ่งเขียน
  `AS PER ATTACHED RIDER` / `AS PER ATTACHED DESCRIPTIONS` จะไปหารายละเอียดใน RIDER ให้
- เตือนเพิ่ม: เลขตู้ผิดหลักตรวจสอบ (ISO 6346), G.W. น้อยกว่า N.W., จำนวนตู้ไม่ตรงกับ TOTAL (IN WORDS),
  เลข B/L ในชื่อไฟล์ไม่ตรงเอกสาร, วันที่ ON BOARD หลังวันออก B/L

**ความหมายของสี:** 🟥 แดง = ไม่ตรง · 🟨 เหลือง = มีข้อมูลฝั่งเดียว · 🟦 ฟ้า = ตรงโดยหาเจอใน RIDER ·
🟩 เขียว = ตรง · ⬜ เทา = ว่างทั้งสองฝั่ง

---

## ใช้ออนไลน์ (ง่ายสุด ไม่ต้องติดตั้ง)

เปิด **<https://rattanao.github.io/B-L-CHECK-REPORT-Export/>** → ลากไฟล์ CLIENT B/L + DRAFT HAL ใส่
(RIDER ที่เป็นไฟล์แยกวางเพิ่มได้) → กด **ตรวจสอบ** → กด **ดาวน์โหลด BL CHECK.xlsx**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่มีการอัปโหลดไปที่ใด ·
คู่มือ: <https://rattanao.github.io/B-L-CHECK-REPORT-Export/guide.html>

---

## วิธีที่ 1 — ใช้แบบดับเบิลคลิก (ไม่ต้องมี Claude)

### ติดตั้งครั้งแรก (ทำครั้งเดียว)

1. ติดตั้ง **Python 3** จาก <https://www.python.org/downloads/>
   - หน้าแรกของตัวติดตั้ง **ติ๊ก ☑ "Add python.exe to PATH"** ก่อนกด Install
2. ดาวน์โหลดโปรแกรมนี้: หน้า repo กดปุ่มเขียว **Code → Download ZIP** แล้วแตกไฟล์ไว้ที่ไหนก็ได้
   - หรือถ้ามี git: `git clone https://github.com/Rattanao/B-L-CHECK-REPORT-Export.git`
3. รันครั้งแรกต้องต่ออินเทอร์เน็ต โปรแกรมจะติดตั้ง `pdfplumber`, `openpyxl` ให้เอง

### ใช้งานทุกครั้ง

1. วางไฟล์ในโฟลเดอร์ `input` — ชื่อไฟล์ต้องมีเลข B/L และไฟล์ของ HAL ต้องมีคำว่า `DRAFT`
   เช่น `CLIENT HASLS22260901995.pdf` + `DRAFT HAL HASLS22260901995.pdf` (ใส่หลาย B/L พร้อมกันได้;
   RIDER แยกไฟล์ให้ชื่อมี `RIDER` เช่น `CLIENT RIDER HASLS22260901995.pdf`)
2. ดับเบิลคลิก **`run-bl-check.bat`** → ผลอยู่ที่ `output\BL CHECK <เลข B/L>.xlsx`
3. หรือดับเบิลคลิก **`open-bl-checker.bat`** เพื่อเปิดหน้าต่างโปรแกรม เลือกไฟล์แล้วกด CHECK

---

## วิธีที่ 2 — ใช้กับ Claude Code

```
git clone https://github.com/Rattanao/B-L-CHECK-REPORT-Export.git %USERPROFILE%\.claude\skills\bl-check
```

เปิด Claude Code ใหม่ ไปที่โฟลเดอร์งาน แล้วพิมพ์ เช่น “เช็ค B/L HASLS22260901995” — รายละเอียดใน [`SKILL.md`](SKILL.md)

---

## โครงสร้าง

| ไฟล์ | หน้าที่ |
|---|---|
| `scripts/bl_checker.py` | โปรแกรมหลัก (หน้าต่าง / command line / ทั้งโฟลเดอร์) |
| `scripts/layouts.json` | ตำแหน่งช่องของแบบฟอร์ม B/L แต่ละแบบ (HAL DRAFT, CLIENT) |
| `docs/index.html`, `docs/core.js` | เวอร์ชันออนไลน์ (ตรรกะเดียวกับ `bl_checker.py`) |
| `guide.html` | คู่มือการใช้งาน |
| `run-bl-check.bat`, `open-bl-checker.bat` | ตัวเรียกใช้บน Windows |

รายงานนี้เป็นตัวช่วยตรวจ ก่อนออก B/L จริงต้องให้คนตรวจทุกช่องที่เป็นสีแดงหรือสีเหลืองเสมอ
