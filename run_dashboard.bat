@echo off
chcp 65001 >nul
cd /d "%~dp0"
set EXCEL_ARABIC_PORT=8765
python web_app.py
if errorlevel 1 pause
