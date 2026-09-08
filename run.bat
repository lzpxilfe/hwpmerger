@echo off
setlocal
cd /d "%~dp0"
echo [HWP Merger] rhwp-structure-actual-a3-template-tab v38 build 2026-09-08
if not defined HWP_RESTART_EVERY_N_TABLE_GROUPS set HWP_RESTART_EVERY_N_TABLE_GROUPS=20
python -c "import pythoncom" >nul 2>nul
if not errorlevel 1 goto run_python
py -c "import pythoncom" >nul 2>nul
if not errorlevel 1 goto run_py
echo.
echo [ERROR] Python 또는 pywin32를 찾지 못했습니다. README의 "처음 실행" 절을 확인하세요.
goto done

:run_python
python hwp_merger_gui.py
goto done

:run_py
py hwp_merger_gui.py

:done
pause
