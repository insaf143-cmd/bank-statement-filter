@echo off
setlocal enabledelayedexpansion
title Push to GitHub - Bank Statement Filter
color 0A

echo ===============================================================================
echo            PUSH PROJECT TO YOUR GITHUB REPOSITORY
echo ===============================================================================
echo.
cd /d "%~dp0"

:: Check if git is installed
where git >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Git is not installed or not in PATH.
    echo Please install Git from https://git-scm.com/
    pause
    exit /b 1
)

:: Prompt user for repo URL
echo Please paste your GitHub repository URL:
echo (Example: https://github.com/your-username/bank-statement-filter.git)
echo.
set /p REPO_URL="Repository URL: "

if "%REPO_URL%"=="" (
    echo [ERROR] No URL entered. Aborting.
    pause
    exit /b 1
)

echo.
echo [*] Initializing git repository (if needed)...
if not exist ".git" (
    git init -b main
)

echo [*] Staging all files...
git add .

echo [*] Committing changes...
git commit -m "feat: Bank Statement Multi-Format Filtering & Inter-Bank Reconciliation Engine"

echo [*] Setting remote origin...
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%

echo [*] Renaming branch to main...
git branch -M main

echo [*] Pushing to GitHub (main)...
git push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===============================================================================
    echo   SUCCESS: Your project has been pushed to GitHub successfully!
    echo ===============================================================================
) else (
    echo.
    echo [!] Push encountered an issue. If your repository already has files (e.g. README),
    echo     trying pull with rebase first:
    git pull origin main --rebase
    git push -u origin main
)

echo.
pause
