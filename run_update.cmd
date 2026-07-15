@echo off
rem lotto-lab auto updater (Task Scheduler: Sat 21:30 lotto / Thu 20:30 pension)
cd /d "%~dp0"
if not exist logs mkdir logs
set PYTHONIOENCODING=utf-8
echo. >> logs\update.log
echo ===== %date% %time% ===== >> logs\update.log
python lotto_lab.py >> logs\update.log 2>&1
echo [exit %errorlevel%] >> logs\update.log

rem push updated data to GitHub mirror (only when origin remote is configured)
git remote get-url origin >nul 2>&1
if not errorlevel 1 (
  git add data report.md >> logs\update.log 2>&1
  git diff --cached --quiet || (
    git commit -m "auto: data update" >> logs\update.log 2>&1
    git push origin main >> logs\update.log 2>&1
    echo [push exit %errorlevel%] >> logs\update.log
  )
)
