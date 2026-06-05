@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
echo Starting shared MySQL (Docker)...
docker compose up -d
if errorlevel 1 (
    echo Failed. Is Docker Desktop running?
    pause
    exit /b 1
)
echo.
echo Shared MySQL is up on port %MYSQL_PORT% (default 3306).
echo Set MYSQL_HOST in .env to this PC's IP for other developers.
docker compose ps
pause
