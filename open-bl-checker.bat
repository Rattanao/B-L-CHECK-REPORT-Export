@echo off
chcp 65001 >nul
cd /d "%~dp0"
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
start "" %PY% scripts\bl_checker.py
