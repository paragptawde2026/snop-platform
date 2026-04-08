@echo off
title SNOP Frontend
echo =============================================
echo   SNOP Frontend (React + Vite)
echo   http://localhost:5173
echo =============================================
cd /d "%~dp0frontend"
echo Starting frontend...
npm run dev
pause
