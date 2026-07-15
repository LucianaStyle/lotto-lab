@echo off
rem lotto-lab auto updater (Task Scheduler: Sat 21:30 lotto / Thu 20:30 pension)
cd /d "%~dp0"
if not exist logs mkdir logs
set PYTHONIOENCODING=utf-8
echo. >> logs\update.log
echo ===== %date% %time% ===== >> logs\update.log
python lotto_lab.py >> logs\update.log 2>&1
echo [exit %errorlevel%] >> logs\update.log
