#!/usr/bin/env bash

# SafeRoad AI - Full-Stack Launcher for Linux / macOS
set -e

echo "=============================================================="
echo "         SafeRoad AI - Intelligent Road Safety Platform       "
echo "=============================================================="
echo ""

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Cleanup background processes on exit
trap 'kill $(jobs -p) 2>/dev/null || true' EXIT

echo "[1/2] Starting FastAPI Backend on http://127.0.0.1:8001 ..."
(cd backend && python3 run_backend.py) &
BACKEND_PID=$!

sleep 2

echo "[2/2] Starting React Vite Frontend on http://localhost:3000 ..."
(cd frontend && npm run dev) &
FRONTEND_PID=$!

echo ""
echo "=============================================================="
echo " SafeRoad AI is running!"
echo " - Frontend: http://localhost:3000"
echo " - Backend API: http://127.0.0.1:8001"
echo " - Interactive Docs: http://127.0.0.1:8001/docs"
echo " Press Ctrl+C to terminate both servers."
echo "=============================================================="
echo ""

wait
