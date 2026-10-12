@echo off
REM Chạy Tool Test Tự Động từ mã nguồn (dành cho người phát triển).
REM Cần Python 3.9+ có: pip install playwright openpyxl pillow   (KHÔNG cần "playwright install")
chcp 65001 >nul
cd /d "%~dp0"
set "PY="
if defined PYTHON_EXE set "PY=%PYTHON_EXE%"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY (where py >nul 2>nul && set "PY=py -3")
if not defined PY set "PY=python"
%PY% -c "import playwright, openpyxl, PIL" 2>nul || (
  echo Dang cai thu vien can thiet...
  %PY% -m pip install playwright openpyxl pillow || (echo [LOI] Khong cai duoc thu vien & pause & exit /b 1)
)
start "" %PY% app.py
