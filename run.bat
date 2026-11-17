@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo   Slide ^& PDF Smart Compressor - Khoi dong May Chu Web
echo   Dia chi: http://127.0.0.1:8080
echo ============================================================
start "" "http://127.0.0.1:8080"
python -m uvicorn app:app --host 127.0.0.1 --port 8080
pause
