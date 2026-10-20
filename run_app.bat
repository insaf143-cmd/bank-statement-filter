@echo off
title Bank Statement Filter & Inter-Bank Reconciliation Engine
cd /d "%~dp0"

echo ======================================================================
echo    Starting Bank Statement Filter & Inter-Bank Reconciliation Engine
echo ======================================================================
echo.

if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8501"
python -m streamlit run app.py --server.port 8501 --server.headless false

pause
