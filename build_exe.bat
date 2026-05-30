@echo off
setlocal

set APP_NAME=Md2Docx
set ENTRY=app.py
set TARGET_PY=3.8

if not exist "%ENTRY%" (
  echo [ERROR] Cannot find %ENTRY% in current folder.
  exit /b 1
)

for /f "tokens=1,2 delims=." %%A in ('python -c "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"') do (
  set PY_MAJ=%%A
  set PY_MIN=%%B
)
set PY_VER=%PY_MAJ%.%PY_MIN%
if not "%PY_VER%"=="%TARGET_PY%" (
  echo [WARN] Current Python is %PY_VER%.
  echo [WARN] For Windows 7 compatibility, build with Python %TARGET_PY% x64.
  echo [WARN] Continue build anyway...
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
