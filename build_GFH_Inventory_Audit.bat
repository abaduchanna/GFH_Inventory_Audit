@echo off
setlocal enabledelayedexpansion
title Build GFH_Inventory_Audit

set "SRCDIR=C:\Users\AbadUmairChanna\Downloads\GitHub\GFH-Inventory-Audit"
set "OUTDIR=C:\Users\AbadUmairChanna\Downloads\GitHub"
set "REPOURL=https://github.com/abaduchanna/GFH-Inventory-Audit.git"
set "WORKBASE=%TEMP%\pyi_build\GFH_Inventory_Audit"

echo.
echo  ============================================================
echo   Building: GFH_Inventory_Audit.exe
echo  ============================================================
echo.

REM Check prerequisites
python --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Python not found in PATH.
    pause
    exit /b 1
)
REM -- Protect sources: compile every local module to native extensions -------
REM  (no Python source and no decompilable bytecode ships inside the EXE)
echo    Protecting sources (compiling to native extensions)...
python -m pip install --quiet cython pyflakes 2>nul
python _protect_build.py GFH_Inventory_Audit.spec
if errorlevel 1 (
    echo    FAILED: source protection (_protect_build.py)
    pause
    exit /b 1
)

python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo    PyInstaller not found. Installing...
    python -m pip install --upgrade pyinstaller
)
git --version >nul 2>&1
if errorlevel 1 (
    echo    ERROR: Git not found in PATH.
    pause
    exit /b 1
)
echo    Prerequisites OK
echo.

REM Clone or force-sync
REM A plain "git pull" breaks when the clone's origin still embeds an old
REM revoked token (401) or when history on GitHub was rewritten. Self-heal
REM the origin URL to the clean public URL (repo is public, no auth needed)
REM and hard-reset to origin/main - this ALWAYS ends up byte-identical to
REM GitHub no matter what happened before.
if exist "%SRCDIR%" (
    echo    Syncing latest from GitHub...
    cd "%SRCDIR%"
    git remote set-url origin "%REPOURL%" >nul 2>&1
    git fetch origin 2>&1
    git reset --hard origin/main 2>&1
) else (
    echo    Cloning GFH-Inventory-Audit...
    git clone "%REPOURL%" "%SRCDIR%" 2>&1
)

cd "%SRCDIR%"
set "BUILD_COMMIT="
for /f "usebackq delims=" %%C in (`git rev-parse --short HEAD 2^>nul`) do set "BUILD_COMMIT=%%C"
echo    Source commit: !BUILD_COMMIT!

REM Clean previous build
echo    Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
del /s /q *.pyc 2>nul

REM Redirect workpath to TEMP
if exist "%WORKBASE%" rmdir /s /q "%WORKBASE%"
mkdir "%WORKBASE%" 2>nul

REM Install deps
if exist "requirements.txt" (
    echo    Installing requirements...
    python -m pip install -r requirements.txt --quiet 2>nul
)

REM Build
echo    Building GFH_Inventory_Audit.spec...
python -m PyInstaller "GFH_Inventory_Audit.spec" --noconfirm --clean --workpath "%WORKBASE%" 2>&1

if errorlevel 1 (
    echo    FAILED: GFH_Inventory_Audit
    pause
    exit /b 1
)

echo    SUCCESS: GFH_Inventory_Audit

REM Copy .exe to output
if exist "dist\GFH_Inventory_Audit.exe" (
    if not exist "%OUTDIR%" mkdir "%OUTDIR%"
    copy /Y "dist\GFH_Inventory_Audit.exe" "%OUTDIR%\GFH_Inventory_Audit.exe" >nul
    if errorlevel 1 (
        echo    WARNING: could not overwrite GFH_Inventory_Audit.exe - close the running exe and rebuild.
    ) else (
        echo    Collected: %OUTDIR%\GFH_Inventory_Audit.exe
    )
) else (
    echo    WARNING: dist\GFH_Inventory_Audit.exe not found
)

echo.
echo  ============================================================
echo   Done: GFH_Inventory_Audit.exe  (source commit !BUILD_COMMIT!)
echo  ============================================================
echo.
pause
endlocal
exit /b 0
