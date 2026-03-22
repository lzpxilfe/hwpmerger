@echo off
cd /d "%~dp0"
pyinstaller --noconfirm HwpMerger.spec
pause
