@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Creating an isolated build environment...
    python -m venv .venv
    if errorlevel 1 goto :error
) else (
    echo [1/3] Reusing the existing build environment...
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 goto :error

echo [2/3] Installing build dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [3/3] Building Aptech Monthly BSR...
python -m PyInstaller --clean --noconfirm AptechMonthlyBSR.spec
if errorlevel 1 goto :error

echo Build completed.
echo Distributable folder: %CD%\dist\AptechMonthlyBSR
echo Launch executable: %CD%\dist\AptechMonthlyBSR\AptechMonthlyBSR.exe
echo Distribute the complete folder, preferably as a ZIP file.
goto :eof

:error
echo Build failed. Review the error output above.
exit /b 1
