@echo off
echo ==========================================
echo  HBS Lab Bill Automation System
echo ==========================================
echo.
echo [1/2] Starting Python backend (port 5001)...
cd backend-dev
start cmd /k "venv\Scripts\activate && python server.py"
timeout /t 3 >nul

echo [2/2] Starting React frontend (port 3000)...
cd ..\frontend-dev
start cmd /k "npm start"

echo.
echo  Backend  : http://localhost:5001
echo  Frontend : http://localhost:3000
echo.
echo Open your browser at http://localhost:3000
pause
