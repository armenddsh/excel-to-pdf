@echo off
cd /d "%~dp0"
echo Running Excel diagnostic...
python diagnose_excel.py
pause
