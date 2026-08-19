@echo off
setlocal
cd /d "%~dp0"
python -c "import pythoncom, PyInstaller" >nul 2>nul
if not errorlevel 1 goto run_python
py -c "import pythoncom, PyInstaller" >nul 2>nul
if not errorlevel 1 goto run_py
echo.
echo [ERROR] pywin32 또는 PyInstaller를 찾지 못했습니다.
echo         python -m pip install pywin32 pyinstaller 를 먼저 실행하세요.
goto done

:run_python
python -m PyInstaller --noconfirm HwpMerger.spec
goto done

:run_py
py -m PyInstaller --noconfirm HwpMerger.spec

:done
pause
