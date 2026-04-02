@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set APP_NAME=MiceTimer
set VERSION_FILE=version.txt
set RELEASES_DIR=releases

echo [INFO] Start build in %cd%

REM ---------- check python ----------
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python not found in PATH.
  goto END
)
python --version

REM ---------- ensure pyinstaller ----------
python -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
  echo [INFO] Installing PyInstaller...
  python -m pip install pyinstaller
  if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    goto END
  )
)

REM ---------- check main.py ----------
if not exist main.py (
  echo [ERROR] main.py not found.
  goto END
)

REM ---------- ensure data directories ----------
if not exist data mkdir data
if not exist data\autosave mkdir data\autosave
if not exist data\export mkdir data\export
if not exist data\templates mkdir data\templates
if not exist %RELEASES_DIR% mkdir %RELEASES_DIR%

REM ---------- version bump ----------
if not exist %VERSION_FILE% echo 1.0.0>%VERSION_FILE%
set /p CUR_VER=<%VERSION_FILE%

for /f "tokens=1,2,3 delims=." %%a in ("%CUR_VER%") do (
  set MAJOR=%%a
  set MINOR=%%b
  set PATCH=%%c
)

if "%MAJOR%"=="" set MAJOR=1
if "%MINOR%"=="" set MINOR=0
if "%PATCH%"=="" set PATCH=0

set /a PATCH=PATCH+1
set NEW_VER=!MAJOR!.!MINOR!.!PATCH!
echo [INFO] Version: %CUR_VER% -^> !NEW_VER!

REM ---------- clean previous build ----------
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist %APP_NAME%.spec del /f /q %APP_NAME%.spec

REM ---------- detect UPX ----------
set "UPX_DIR="
if exist "upx\upx.exe" (
  set "UPX_DIR=upx"
  echo [INFO] UPX found at upx\upx.exe
) else (
  echo [INFO] UPX not found. Build without UPX.
)

REM ---------- build (onefile, size-optimized) ----------
echo [INFO] Building onefile optimized package...

if defined UPX_DIR (
  python -m PyInstaller ^
    --noconfirm ^
    --windowed ^
    --onefile ^
    --name %APP_NAME% ^
    --add-data "data;data" ^
    --exclude-module tkinter ^
    --exclude-module unittest ^
    --exclude-module test ^
    --exclude-module matplotlib ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module scipy ^
    --upx-dir "!UPX_DIR!" ^
    main.py > build_log.txt 2>&1
) else (
  python -m PyInstaller ^
    --noconfirm ^
    --windowed ^
    --onefile ^
    --name %APP_NAME% ^
    --add-data "data;data" ^
    --exclude-module tkinter ^
    --exclude-module unittest ^
    --exclude-module test ^
    --exclude-module matplotlib ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module scipy ^
    main.py > build_log.txt 2>&1
)

if errorlevel 1 (
  echo [ERROR] Build failed. Check build_log.txt
  powershell -Command "Get-Content build_log.txt -Tail 80"
  goto END
)

REM ---------- package release ----------
set "OUT_DIR=%RELEASES_DIR%\%APP_NAME%_v!NEW_VER!"
if exist "!OUT_DIR!" rmdir /s /q "!OUT_DIR!"
mkdir "!OUT_DIR!" >nul 2>nul

copy /y "dist\%APP_NAME%.exe" "!OUT_DIR!\%APP_NAME%.exe" >nul
if errorlevel 1 (
  echo [ERROR] Failed to copy exe to release directory.
  goto END
)

(
echo %APP_NAME% v!NEW_VER!
echo.
echo Portable onefile build.
echo Run: %APP_NAME%.exe
echo.
echo Notes:
echo - First launch may be slower due to onefile unpacking.
echo - If antivirus blocks it, add an exclusion rule.
echo - Export files are saved to: ^<app dir^>\data\export\
) > "!OUT_DIR!\README.txt"

REM ---------- update version file ----------
echo !NEW_VER!>%VERSION_FILE%

REM ---------- report ----------
for %%I in ("!OUT_DIR!\%APP_NAME%.exe") do set EXE_SIZE=%%~zI
echo [SUCCESS] Output:  !OUT_DIR!\%APP_NAME%.exe
echo [SUCCESS] Size:    !EXE_SIZE! bytes
echo [SUCCESS] Version: !NEW_VER!

:END
echo.
pause
