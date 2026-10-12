@echo off
chcp 65001 >nul
title Tool Test Tu Dong
rem =====================================================================================================
rem  Chạy Tool Test Tự Động từ mã nguồn — bấm đúp file này.
rem  Lần đầu: tự cài Python (nếu máy chưa có) + thư viện vào thư mục .venv bên cạnh (mất 1–3 phút, cần mạng).
rem  Các lần sau: mở tool ngay.
rem =====================================================================================================
cd /d "%~dp0"
set "VPY=%~dp0.venv\Scripts\python.exe"
set "VPYW=%~dp0.venv\Scripts\pythonw.exe"

if exist "%VPY%" (
  "%VPY%" -c "import playwright, openpyxl, PIL" >nul 2>nul && goto run
)

echo [1/3] Tim Python tren may...
set "PY="
for %%V in (313 312 311 310) do (
  if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
)
if not defined PY (py -3 -c "import sys; assert sys.version_info >= (3, 9)" >nul 2>nul && set "PY=py -3")
if not defined PY (python -c "import sys; assert sys.version_info >= (3, 9)" >nul 2>nul && set "PY=python")
if not defined PY (
  echo     Chua co Python. Dang cai Python 3.12 bang winget...
  where winget >nul 2>nul || goto nopython
  winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements || goto nopython
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY goto nopython
echo     Dung: %PY%

echo [2/3] Tao moi truong rieng .venv ...
%PY% -m venv "%~dp0.venv" || (echo [LOI] Khong tao duoc .venv & pause & exit /b 1)

echo [3/3] Cai thu vien (playwright, openpyxl, pillow) - khoang 45 MB, vui long doi...
"%VPY%" -m pip install --disable-pip-version-check -q -r requirements.txt || (
  echo [LOI] Khong cai duoc thu vien. Kiem tra mang / proxy cong ty roi chay lai file nay.
  pause & exit /b 1
)
echo     Xong. Khong can "playwright install": tool dung Edge / Chrome co san tren may.

:run
echo Dang mo Tool Test Tu Dong...
start "" "%VPYW%" app.py
exit /b 0

:nopython
echo.
echo [LOI] Khong tu cai duoc Python.
echo  - Cach 1: tai Python 3.12 tai https://www.python.org/downloads/ , khi cai TICK "Add python.exe to PATH",
echo           cai xong bam dup lai file ChayTool.bat nay.
echo  - Cach 2: tai ban EXE (khong can Python) o muc Releases cua trang GitHub.
start "" https://www.python.org/downloads/
pause
exit /b 1
