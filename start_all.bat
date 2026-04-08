@echo off
echo =============================================
echo   Starting SNOP Platform (No Docker)
echo =============================================
echo.

REM Check PostgreSQL on port 5433
netstat -ano | findstr ":5433" | findstr "LISTENING" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PostgreSQL is NOT running on port 5433.
    echo Please start PostgreSQL 18 service before running this script.
    echo   - Open Services (services.msc)
    echo   - Start "postgresql-x64-18" service
    pause
    exit /b 1
)
echo [OK] PostgreSQL is running on port 5433

REM Start backend in a new window
echo Starting Backend...
start "SNOP Backend" cmd /k "cd /d "%~dp0backend" && call venv\Scripts\activate && python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload"

REM Wait 3 seconds for backend to initialise
timeout /t 3 /nobreak >nul

REM Start frontend in a new window
echo Starting Frontend...
start "SNOP Frontend" cmd /k "cd /d "%~dp0frontend" && npm run dev"

echo.
echo =============================================
echo   SNOP is starting up!
echo   Backend  : http://localhost:8000
echo   Frontend : http://localhost:5173
echo   API Docs : http://localhost:8000/api/docs
echo =============================================
echo.
echo Both windows will open separately.
echo Close those windows to stop the servers.
pause
