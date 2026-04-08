@echo off
title SNOP Backend
echo =============================================
echo   SNOP Backend (FastAPI + Uvicorn)
echo   http://localhost:8000
echo   API Docs: http://localhost:8000/api/docs
echo =============================================
cd /d "%~dp0backend"
call venv\Scripts\activate
echo Starting backend...
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
pause
