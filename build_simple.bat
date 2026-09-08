@echo off
REM ============================================================
REM   TK PH Converter - SIMPLE build (just double-click)
REM   No admin check, no chcp, no fanciness - just works.
REM ============================================================

REM Locate tk-ph-converter folder automatically.
REM This file should be inside the tk-ph-converter folder.
cd /d "%~dp0"

set "APP_NAME=TK-PH-Converter"

echo.
echo ============================================================
echo  TK PH Converter - Simple Build
echo  Current dir: %CD%
echo ============================================================

REM --- 1. Detect or download Python ---
where python >nul 2>&1
if not errorlevel 1 (
    echo [OK] Python found:
    python --version
    goto :have_python
)

where py >nul 2>&1
if not errorlevel 1 (
    echo [OK] Python (py launcher) found:
    py -3 --version
    goto :have_python_py
)

echo [INFO] Python not detected. Downloading Python 3.12 installer...
echo        (This downloads ~25 MB and takes 1-2 minutes)
echo.

set "PY_INSTALLER=%TEMP%\python-installer.exe"
powershell -Command "$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PY_INSTALLER%'"
if errorlevel 1 (
    echo [ERROR] Download failed. Check your internet.
    pause
    exit /b 1
)

echo [INFO] Installing Python (this takes 1-2 minutes)...
"%PY_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_doc=0 Include_launcher=1 InstallLauncherAllUsers=1 TargetDir=C:\Python312
if errorlevel 1 (
    echo [ERROR] Python install failed.
    pause
    exit /b 1
)

set "PATH=C:\Python312;C:\Python312\Scripts;%PATH%"
echo [OK] Python installed.
goto :have_python

:have_python_py
set "PATH_FOR_PY="
goto :have_python

:have_python

REM --- 2. Create venv ---
if not exist ".venv\Scripts\python.exe" (
    echo.
    echo [INFO] Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        py -3 -m venv .venv
        if errorlevel 1 (
            echo [ERROR] Failed to create venv.
            pause
            exit /b 1
        )
    )
)

set "PY=.venv\Scripts\python.exe"
echo [OK] Using: %PY%

REM --- 3. Install dependencies ---
echo.
echo [INFO] Installing dependencies (1-2 minutes)...
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r requirements.txt
"%PY%" -m pip install pyinstaller
if errorlevel 1 (
    echo [ERROR] Dependency install failed.
    pause
    exit /b 1
)

REM --- 4. Build with PyInstaller ---
echo.
echo [INFO] Building exe (3-5 minutes, please wait)...
if exist "build" rmdir /S /Q "build"
if exist "dist" rmdir /S /Q "dist"
"%PY%" -m PyInstaller --noconfirm "%APP_NAME%.spec"
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

REM --- 5. Copy template ---
if not exist "dist\assets" mkdir "dist\assets"
copy /Y "assets\batch-product-source.xlsx" "dist\assets\" >nul

echo.
echo ============================================================
echo   BUILD COMPLETE!
echo.
echo   Output:  %CD%\dist\%APP_NAME%.exe
echo.
echo   Go to: %CD%\dist
echo   Double-click %APP_NAME%.exe to run.
echo ============================================================
echo.
pause
