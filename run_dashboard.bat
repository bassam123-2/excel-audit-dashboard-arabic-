@echo off
chcp 65001 >nul
cd /d "%~dp0"
set EXCEL_ARABIC_PORT=8765
python manage.py runserver 0.0.0.0:%EXCEL_ARABIC_PORT%
if errorlevel 1 pause
