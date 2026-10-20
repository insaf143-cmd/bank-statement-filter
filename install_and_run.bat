@echo off
setlocal enabledelayedexpansion
title Bank Statement Filter - Automated 1-Click Setup & Launch
color 0B

echo ===============================================================================
echo     BANK STATEMENT MULTI-FORMAT FILTER & INTER-BANK RECONCILIATION ENGINE
echo ===============================================================================
echo.
echo [*] Checking system environment and dependencies...

cd /d "%~dp0"

:: 1. Check Python installation
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Python was not found on your system!
    echo Please install Python 3.9 or newer from https://www.python.org/downloads/
    echo Make sure to check the box "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PY_VER=%%i
echo [*] Detected: %PY_VER%

:: 2. Check or create virtual environment
if not exist ".venv" (
    echo.
    echo [*] Creating isolated virtual environment in .venv...
    python -m venv .venv
    if %ERRORLEVEL% neq 0 (
        echo [!] Virtual environment creation failed. Using global Python...
        goto install_global
    )
)

echo [*] Activating virtual environment...
call .venv\Scripts\activate.bat

echo [*] Upgrading pip...
python -m pip install --upgrade pip --quiet

echo [*] Installing required dependencies from requirements.txt...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to install dependencies. Please check your internet connection.
    pause
    exit /b 1
)
goto launch_app

:install_global
echo [*] Installing dependencies directly into global Python...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt

:launch_app
echo.
echo ===============================================================================
echo   SUCCESS: Installation complete! Launching application...
echo   URL: http://localhost:8501
echo ===============================================================================
echo.

:: Automatically open browser after 2 seconds in parallel
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8501"

:: Run Streamlit
python -m streamlit run app.py --server.port 8501 --server.headless false

pause
