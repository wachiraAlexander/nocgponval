@echo off
REM GPON Application - Windows Fix and Run Script
REM This script clears the ChromeDriver cache and fixes Windows compatibility issues

echo.
echo ========================================
echo GPON Application - Windows Troubleshooter
echo ========================================
echo.

REM Check if Python is installed
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://www.python.org/
    pause
    exit /b 1
)

REM Run the ChromeDriver fix script
echo Running ChromeDriver cache cleaner...
echo.
python fix_chromedriver.py

if %errorlevel% neq 0 (
    echo.
    echo Error: Fix script failed. Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Now starting GPON Application...
echo ========================================
echo.

REM Run the application
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application

pause
