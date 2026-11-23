@echo off
setlocal
cd /d "%~dp0"
echo [1/3] Installing build dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
echo [2/3] Building Aptech Monthly BSR...
python -m PyInstaller --clean --noconfirm AptechMonthlyBSR.spec
if errorlevel 1 goto :error
echo [3/3] Build completed.
echo Distributable folder: %CD%\dist\AptechMonthlyBSR
echo Launch executable: %CD%\dist\AptechMonthlyBSR\AptechMonthlyBSR.exe
echo Distribute the complete folder, preferably as a ZIP file.
goto :eof
:error
echo Build failed. Review the error output above.
exit /b 1
