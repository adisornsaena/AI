@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title ผู้ช่วย AI สำนักงาน

rem ตัวแปรที่ตั้งก่อนรันได้: API_KEY (คีย์), PORT (ค่าเริ่มต้น 5000), NO_BROWSER=1 (ไม่เปิดเบราว์เซอร์อัตโนมัติ)

rem --- 1) หา Python (ลอง python ก่อน แล้วลอง py -3) ---
set "PY=python"
%PY% --version >nul 2>nul
if errorlevel 1 set "PY=py -3"
%PY% --version >nul 2>nul
if errorlevel 1 goto nopython

%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if errorlevel 1 goto oldpython

rem --- 2) ติดตั้งไลบรารี ถ้ายังไม่มี (ครั้งแรกเท่านั้น) ---
%PY% -c "import flask, openai, docx, pypdf" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่จำเป็น ครั้งแรกอาจใช้เวลาสักครู่...
  %PY% -m pip install -r requirements.txt
  if errorlevel 1 goto pipfail
)

rem --- 3) ขอ API Key ถ้ายังไม่ได้ตั้งไว้ (ใช้เฉพาะหน้าต่างนี้ ไม่บันทึกลงไฟล์) ---
if not defined API_KEY (
  echo.
  echo ยังไม่ได้ตั้งค่า API Key
  echo วางคีย์แล้วกด Enter  คีย์จะแสดงบนหน้าจอ และใช้เฉพาะรอบนี้ ไม่ถูกบันทึกลงไฟล์
  set /p "API_KEY=API Key: "
)
if not defined API_KEY goto nokey

if not defined PORT set "PORT=5000"

echo.
echo เปิดเซิร์ฟเวอร์ที่ http://127.0.0.1:%PORT%
echo ปิดโปรแกรมโดยกด Ctrl+C หรือปิดหน้าต่างนี้
echo.

rem รอเซิร์ฟเวอร์เริ่มทำงาน 3 วินาทีแล้วเปิดเบราว์เซอร์
if not defined NO_BROWSER start "" cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:%PORT%"

%PY% app.py
echo.
echo เซิร์ฟเวอร์หยุดทำงานแล้ว หากมีข้อความผิดพลาดด้านบน โปรดอ่านก่อนปิดหน้าต่าง
pause
exit /b 0

:nopython
echo ไม่พบ Python บนเครื่องนี้
echo โปรดติดตั้งจาก https://www.python.org/downloads/ และติ๊ก "Add Python to PATH" ตอนติดตั้ง แล้วดับเบิลคลิกไฟล์นี้อีกครั้ง
pause
exit /b 1

:oldpython
echo Python เวอร์ชันเก่าเกินไป ต้องใช้ 3.10 ขึ้นไป  เวอร์ชันที่พบ:
%PY% --version
echo โปรดติดตั้งเวอร์ชันใหม่จาก https://www.python.org/downloads/
pause
exit /b 1

:pipfail
echo ติดตั้งไลบรารีไม่สำเร็จ ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต แล้วลองอีกครั้ง
pause
exit /b 1

:nokey
echo ไม่ได้ใส่ API Key จึงไม่สามารถเริ่มได้
pause
exit /b 1
