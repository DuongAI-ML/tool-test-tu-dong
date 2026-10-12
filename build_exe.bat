@echo off
chcp 65001 >nul
rem ===== Đóng gói ToolTestTuDong.exe bằng PyInstaller (1 thư mục, không cửa sổ đen, có icon) =====
rem  Dạng THƯ MỤC (onedir) mở nhanh hơn và ít bị phần mềm diệt virus báo nhầm hơn dạng 1 file.
rem  Không kèm trình duyệt: tool dùng Edge / Chrome có sẵn trên máy người dùng.
cd /d "%~dp0"
set "PY=%TOOLTEST_PYTHON%"
if "%PY%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if "%PY%"=="" set "PY=python"
"%PY%" -m PyInstaller --version >nul 2>nul || (
  echo Chua co PyInstaller. Cai bang: "%PY%" -m pip install pyinstaller
  pause & exit /b 1
)
"%PY%" -c "import playwright, openpyxl, PIL" || (
  echo Chua du thu vien. Cai bang: "%PY%" -m pip install -r requirements.txt
  pause & exit /b 1
)
"%PY%" tools\make_icon.py || (pause & exit /b 1)
"%PY%" -m PyInstaller --noconfirm --clean --onedir --noconsole ^
  --name ToolTestTuDong --icon "%~dp0web\app.ico" ^
  --add-data "%~dp0web;web" ^
  --add-data "%~dp0demo;demo" ^
  --collect-all playwright ^
  --exclude-module numpy --exclude-module pandas --exclude-module matplotlib --exclude-module scipy ^
  --exclude-module pytest --exclude-module IPython --exclude-module PyQt5 --exclude-module PySide6 ^
  --distpath dist --workpath build --specpath build ^
  app.py || (echo DONG GOI THAT BAI & pause & exit /b 1)
xcopy /e /i /y vi-du dist\ToolTestTuDong\vi-du >nul
copy /y "HUONG-DAN-SU-DUNG.txt" dist\ToolTestTuDong\ >nul 2>nul
"%PY%" tools\make_zip.py || (pause & exit /b 1)
echo.
echo Xong: dist\ToolTestTuDong\ToolTestTuDong.exe  va file nen de gui: dist\ToolTestTuDong_*.zip
if not "%1"=="--no-pause" pause
