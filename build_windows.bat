@echo off
chcp 65001 >nul
REM ===================================================================
REM   TK PH Converter - Windows one-click build
REM
REM   IMPORTANT: Right-click this file and choose
REM   "Run as administrator" (or it may fail to install Python).
REM ===================================================================
setlocal EnableDelayedExpansion

set "APP_NAME=TK-PH-Converter"
set "ROOT=%~dp0"
pushd "%ROOT%"

echo.
echo ============================================================
echo   TK PH Converter - Build Script
echo   Working dir: %ROOT%
echo ============================================================

echo.
echo === [1/6] Checking Python ===
set "PY_CMD="

where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys; print('  Python (py):', sys.version.split()[0])" 2>nul
    if not errorlevel 1 set "PY_CMD=py -3" & goto :have_py
)

where python >nul 2>&1
if not errorlevel 1 (
    python -c "import sys; print('  Python (python):', sys.version.split()[0])" 2>nul
    if not errorlevel 1 set "PY_CMD=python" & goto :have_py
)

REM No Python found - auto-install
echo   [i] Python not detected. Installing Python 3.12...
echo       (Downloads ~25 MB, takes 1-2 minutes)
echo.

where winget >nul 2>&1
if not errorlevel 1 (
    echo   Using winget to install Python...
    winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 goto :manual_install
) else (
:manual_install
    echo   Downloading Python 3.12 installer...
    set "PY_INSTALLER=%TEMP%\python-3.12.7-amd64.exe"
    powershell -Command "$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PY_INSTALLER%'"
    if errorlevel 1 (
        echo [X] Python download failed.
        pause
        popd
        exit /b 1
    )
    echo   Installing Python silently (1-2 minutes)...
    "%PY_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_doc=0 Include_launcher=1 InstallLauncherAllUsers=1 TargetDir=C:\Python312
    if errorlevel 1 (
        echo [X] Python installation failed.
        pause
        popd
        exit /b 1
    )
    set "PY_CMD=C:\Python312\python.exe"
    set "PATH=C:\Python312;C:\Python312\Scripts;%PATH%"
)

:have_py
echo   Using: %PY_CMD%

echo.
echo === [2/6] Creating venv ===
if not exist ".venv\Scripts\python.exe" (
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo [X] Failed to create venv.
        pause
        popd
        exit /b 1
    )
    echo   Created .venv
) else (
    echo   Reusing existing .venv
)
set "PY=.venv\Scripts\python.exe"

echo.
echo === [3/6] Upgrading pip ===
"%PY%" -m pip install --upgrade pip --quiet

echo.
echo === [4/6] Installing dependencies ===
"%PY%" -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [X] Dependency install failed.
    pause
    popd
    exit /b 1
)
"%PY%" -m pip install pyinstaller --quiet
if errorlevel 1 (
    echo [X] PyInstaller install failed
    pause
    popd
    exit /b 1
)
echo   Dependencies ready

echo.
echo === [5/6] Building with PyInstaller (3-5 minutes) ===
if exist "build"   rmdir /S /Q "build"
if exist "dist"    rmdir /S /Q "dist"
"%PY%" -m PyInstaller --noconfirm "%APP_NAME%.spec"
if errorlevel 1 (
    echo.
    echo [X] PyInstaller build failed. Please screenshot the error above.
    pause
    popd
    exit /b 1
)

echo.
echo === [6/6] Copying template to dist ===
if not exist "dist\assets"   mkdir "dist\assets"
copy /Y "assets\batch-product-source.xlsx" "dist\assets\" >nul
echo   Copied assets\batch-product-source.xlsx

echo.
echo ============================================================
echo   BUILD COMPLETE!
echo.
echo   Output dir:  %ROOT%dist
echo   Executable:  dist\%APP_NAME%.exe
echo   Template:    dist\assets\batch-product-source.xlsx
echo.
echo   Next steps:
echo     1. Open the dist\ folder
echo     2. Double-click %APP_NAME%.exe to run
echo     3. To distribute: zip the entire dist\ folder
echo ============================================================
echo.
pause
popd
endlocal
