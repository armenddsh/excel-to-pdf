@echo off
cd /d "%~dp0"
echo Running Excel conversion test...
python test_conversion.py
pause
