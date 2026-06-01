@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Building Excel Arabic Dashboard...
python -m pip install -r requirements-build.txt
if errorlevel 1 (
    echo pip install failed
    pause
    exit /b 1
)
python -m PyInstaller excel_arabic_dashboard.spec --noconfirm
if errorlevel 1 (
    echo Build failed
    pause
    exit /b 1
)
echo.
echo Done: dist\excel_arabic_dashboard.exe
echo Optional: copy .env or smtp.env next to the exe for email settings.
pause
