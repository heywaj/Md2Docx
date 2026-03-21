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

set "PANDOC_PATH="
for /f "delims=" %%I in ('where pandoc 2^>nul') do (
  set "PANDOC_PATH=%%I"
  goto :pandoc_found
)

echo [ERROR] pandoc not found.
echo For build machine, install pandoc or make it available in PATH.
exit /b 1

:pandoc_found
echo [INFO] Bundling pandoc from: %PANDOC_PATH%

pyinstaller --noconfirm --clean --onefile --windowed --name %APP_NAME% --add-binary "%PANDOC_PATH%;." %ENTRY%
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo.
echo [OK] Build finished: dist\%APP_NAME%.exe
endlocal
