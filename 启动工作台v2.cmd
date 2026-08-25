@echo off
set "PYTHON_EXE=C:\Users\12616\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON_EXE%" (
  echo Python runtime not found.
  pause
  exit /b 1
)
title Workbench v2.0
echo ========================================
echo   Teaching Archive Workbench v2.0
echo   FastAPI + Vue 3
echo ========================================
echo.

cd /d "%~dp0"

echo [1/2] Checking frontend build...
if not exist "frontend\dist\index.html" (
    echo   Building frontend...
    cd frontend
    call npm install
    call npm run build
    cd ..
    echo   Frontend build done.
) else (
    echo   Frontend already built.
)

echo.
echo [2/2] Starting server...
echo   URL: http://127.0.0.1:8080
echo   Press Ctrl+C to stop
echo.

start "" /b cmd /c "timeout /t 3 /nobreak >nul && start "" http://127.0.0.1:8080"

"%PYTHON_EXE%" "%~dp0api.py"

pause
