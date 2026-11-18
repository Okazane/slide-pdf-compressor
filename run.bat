@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Slide ^& PDF Smart Compressor
cd /d "%~dp0"

echo ==============================================================
echo   Slide ^& PDF Smart Compressor
echo   He thong Nen ^& Giai nen PDF / PowerPoint Thong Minh
echo ==============================================================
echo.

:: 1. Tu dong tim Python (.venv -> venv -> python -> py -3)
set "PY_CMD="
if exist ".venv\Scripts\python.exe" (
    set "PY_CMD=.venv\Scripts\python.exe"
    goto :found_py
)
if exist "venv\Scripts\python.exe" (
    set "PY_CMD=venv\Scripts\python.exe"
    goto :found_py
)
where python >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :found_py
)
where py >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3"
    goto :found_py
)

echo [LOI] Khong tim thay Python tren may tinh!
echo Vui long cai dat Python 3.10+ tai: https://www.python.org/downloads/
echo (Luu y tich chon "Add Python to PATH" khi cai dat).
echo.
if /i "%~1"=="--check" exit /b 1
pause
exit /b 1

:found_py
:: 2. Kiem tra thu vien va tu dong cai dat neu thieu
%PY_CMD% -c "import fastapi, uvicorn, multipart, pypdf, pypdfium2, PIL, pptx" >nul 2>&1
if errorlevel 1 (
    echo [*] Phat hien thieu thu vien Python. Dang tu dong cai dat tu requirements.txt...
    %PY_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [LOI] Cai dat thu vien that bai. Vui long kiem tra ket noi mang!
        if /i "%~1"=="--check" exit /b 1
        pause
        exit /b 1
    )
    echo [OK] Da cai dat day du thu vien!
    echo.
)

:: 3. Kiem tra trang thai cong 8080 (neu app dang chay thi mo trinh duyet ngay, neu ban thi tu chon cong trong)
set "PORT_STATE=FREE"
set "APP_PORT=8080"
for /f "tokens=1,2" %%A in ('%PY_CMD% -c "import socket,urllib.request;h='';exec('try:\n h=urllib.request.urlopen(\"http://127.0.0.1:8080/\",timeout=1).read().decode(\"utf-8\",\"ignore\")\nexcept Exception: pass');print('RUNNING 8080' if 'Smart Compressor' in h else 'FREE '+str(next((p for p in range(8080,8095) if socket.socket().connect_ex(('127.0.0.1',p))!=0),8080)))"') do (
    set "PORT_STATE=%%A"
    set "APP_PORT=%%B"
)

set "APP_URL=http://127.0.0.1:%APP_PORT%"

if /i "%~1"=="--check" (
    echo [CHECK OK] Python="%PY_CMD%" State=%PORT_STATE% URL=%APP_URL%
    exit /b 0
)

if /i "%PORT_STATE%"=="RUNNING" (
    echo [OK] May chu da chay san tai: %APP_URL%
    echo [*] Dang mo trinh duyet...
    start "" "%APP_URL%"
    timeout /t 2 >nul
    exit /b 0
)

echo [*] Dia chi truy cap : %APP_URL%
echo [*] Dung luong ho tro: Toi da 500 MB (.PDF, .PPTX, .PPT)
echo [*] Nhan Ctrl + C hoac dong cua so nay khi muon tat may chu.
echo ==============================================================
echo.

:: 4. Doi may chu san sang roi tu dong mo trinh duyet (tranh loi Connection Refused)
start /b "" powershell -NoProfile -WindowStyle Hidden -Command "for ($i=0; $i -lt 40; $i++) { try { $r = Invoke-WebRequest -Uri '%APP_URL%/' -UseBasicParsing -TimeoutSec 1; if ($r.StatusCode -eq 200) { Start-Process '%APP_URL%'; break } } catch {} Start-Sleep -Milliseconds 300 }"

:: 5. Khoi chay may chu Uvicorn
%PY_CMD% -m uvicorn app:app --host 127.0.0.1 --port %APP_PORT%

echo.
echo [!] May chu da dung hoat dong.
pause
