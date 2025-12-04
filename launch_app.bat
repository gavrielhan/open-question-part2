@echo off
REM Launcher script for Topic Generator Web App (Windows)
REM This script starts the Flask web application

REM Get the directory where this script is located
cd /d "%~dp0"

REM Activate virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

REM Check if .env file exists
if not exist ".env" (
    echo.
    echo ERROR: .env file not found!
    echo.
    echo Please create a .env file with your API configuration in:
    echo %CD%
    echo.
    pause
    exit /b 1
)

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.7 or higher.
    echo.
    pause
    exit /b 1
)

REM Start the web application
echo.
echo ========================================
echo   Topic Generator Web App
echo ========================================
echo.
echo Starting server...
echo Your browser will open automatically.
echo.
echo Press Ctrl+C to stop the server.
echo.

python topic_generator_app.py

REM Keep window open if there's an error
if errorlevel 1 (
    echo.
    echo.
    echo An error occurred. Check the messages above.
    echo.
    pause
)

