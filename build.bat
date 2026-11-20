@echo off
rem ============================================================
rem  OFD Reader / Converter - one-file build script
rem  Target: Windows 7 SP1 or later, 64-bit
rem  Toolchain: Python 3.8 (last CPython supporting Windows 7)
rem ============================================================
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PY=%LOCALAPPDATA%\Programs\Python\Python38\python.exe"
if not exist "%PY%" set "PY=python"

echo ==== [1/3] Environment ====
"%PY%" -c "import sys;print('Python',sys.version.split()[0],'(64-bit)' if sys.maxsize>2**32 else '(32-bit)')"
"%PY%" -c "import docx,lxml.etree,PIL,PyInstaller;print('deps: python-docx / lxml / Pillow / PyInstaller  OK')"
if errorlevel 1 goto :deperror

echo ==== [2/3] PyInstaller (spec mode) ====
"%PY%" -m PyInstaller ofd_reader.spec --noconfirm --clean
if errorlevel 1 goto :builderror

echo ==== [3/3] Result ====
for %%F in ("dist\*.exe") do echo   %%~fF    (%%~zF bytes)
echo.
echo Build OK. Copy the exe anywhere and run it.
echo To add it to the right-click "Open with" list:
echo   run the exe, then menu Tools - Add to "Open with"
exit /b 0

:deperror
echo.
echo [ERROR] Missing dependencies. Install them with:
echo   "%PY%" -m pip install -r requirements.txt pyinstaller
exit /b 1

:builderror
echo.
echo [ERROR] Build failed. See the messages above.
exit /b 2
