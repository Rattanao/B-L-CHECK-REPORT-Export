@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===== B-L-CHECK =====
echo วางไฟล์ CLIENT B/L และ DRAFT HAL (.pdf) ไว้ในโฟลเดอร์ input ก่อน (ชื่อไฟล์มีเลข B/L, ไฟล์ HAL มีคำว่า DRAFT)
echo.

rem หา Python ที่รันได้จริง (ตัว python.exe ของ Microsoft Store บางเครื่องเป็นแค่ทางลัดที่รันไม่ได้)
set PY=
py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY python -c "import sys" >nul 2>nul && set PY=python
if not defined PY (
  echo !! ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ แล้วติ๊ก "Add python.exe to PATH"
  pause
  exit /b 1
)

%PY% -c "import pdfplumber, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ครั้งแรกครั้งเดียว...
  %PY% -m pip install -r requirements.txt
)

if not exist output mkdir output
%PY% scripts\bl_checker.py --cli input --outdir output
echo.
echo ผลลัพธ์อยู่ที่ output\BL CHECK ^<เลข B/L^>.xlsx  (ถ้าเปิดไฟล์ค้างไว้ใน Excel ให้ปิดก่อนรัน)
pause
