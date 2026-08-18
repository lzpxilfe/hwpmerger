@echo off
chcp 65001 > nul
cd /d "%~dp0"

if exist do_git_push.py del /f /q do_git_push.py
if exist local_commit.py del /f /q local_commit.py
if exist test_split.py del /f /q test_split.py

echo [1/3] Synchronizing with remote...
git pull origin master --rebase

echo [2/3] Staging and Committing...
git add -A
git commit -m "feat: HWP table auto-split and FileSaveBlock extraction"

echo [3/3] Pushing to GitHub...
git push origin master

echo.
echo ========================================================
echo Git Push Completed Successfully!
echo ========================================================
pause
