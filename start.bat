@echo off
title SafeRoad AI - Full-Stack Launcher
color 0A

echo ==============================================================
echo            SafeRoad AI - Intelligent Road Safety Platform
echo ==============================================================
echo.

cd /d "%~dp0"

echo [1/2] Launching FastAPI Backend on http://127.0.0.1:8001 ...
start "SafeRoad AI - Backend" cmd /k "cd backend && python run_backend.py"

timeout /t 2 >nul

echo [2/2] Launching React Vite Frontend on http://localhost:3000 ...
start "SafeRoad AI - Frontend" cmd /k "cd frontend && npm run dev"

echo.
echo ==============================================================
echo  SafeRoad AI is now running!
echo  - Frontend: http://localhost:3000
echo  - Backend API: http://127.0.0.1:8001
echo  - API Docs: http://127.0.0.1:8001/docs
echo ==============================================================
echo.
