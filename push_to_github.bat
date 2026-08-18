@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo [1/3] Staging all files...
git add -A
git commit -m "feat: HWP table auto-split and FileSaveBlock extraction"

echo [2/3] Rebasing with remote...
git pull origin master --rebase

echo [3/3] Pushing to GitHub...
git push origin master

echo.
echo ========================================================
echo Done! Git Push Completed.
echo ========================================================
pause
