@echo off
setlocal

set APP_NAME=Md2Docx
set ENTRY=app.py

if not exist "%ENTRY%" (
  echo [ERROR] Cannot find %ENTRY% in current folder.
  exit /b 1
)

where pyinstaller >nul 2>&1
if errorlevel 1 (
  echo [ERROR] PyInstaller not found.
  echo Install it with: pip install -r requirements-build.txt
  exit /b 1
)

pyinstaller --noconfirm --clean --onefile --windowed --name %APP_NAME% %ENTRY%
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo.
echo [OK] Build finished: dist\%APP_NAME%.exe
endlocal
