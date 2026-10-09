@echo off
setlocal
cd /d "%~dp0"
echo ==============================================
echo Master Report Center v2.0.2 Release Build
echo ==============================================
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
python -m PyInstaller --noconfirm --clean "Golden Fir Report Center.spec"
if errorlevel 1 (
  echo.
  echo BUILD FAILED
  pause
  exit /b 1
)
echo.
echo BUILD COMPLETE
echo Output: dist\Master Report Center\
echo Test this folder on a clean Windows PC before public release.
pause
