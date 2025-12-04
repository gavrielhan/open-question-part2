@echo off
REM Automated setup script for Windows
REM This script helps set up the Topic Generator application

echo.
echo ========================================
echo   Topic Generator - Windows Setup
echo ========================================
echo.

REM Get the directory where this script is located
cd /d "%~dp0"

REM Check if Python is installed
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH!
    echo.
    echo Please install Python 3.9 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

python --version
echo ✓ Python is installed
echo.

REM Check Python version
echo [2/5] Checking Python version...
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Python version: %PYTHON_VERSION%
echo.

REM Create virtual environment
echo [3/5] Creating virtual environment...
if exist "venv" (
    echo Virtual environment already exists. Skipping...
) else (
    python -m venv venv
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to create virtual environment!
        echo.
        pause
        exit /b 1
    )
    echo ✓ Virtual environment created
)
echo.

REM Activate virtual environment and install dependencies
echo [4/5] Installing dependencies...
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
    pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies!
        echo.
        pause
        exit /b 1
    )
    echo ✓ Dependencies installed
) else (
    echo ERROR: Virtual environment not found!
    pause
    exit /b 1
)
echo.

REM Check for .env file
echo [5/5] Checking configuration...
if exist ".env" (
    echo ✓ .env file exists
    echo.
    echo NOTE: Make sure your .env file contains valid API credentials
) else (
    echo ⚠ .env file not found!
    echo.
    if exist "env_example.txt" (
        echo Creating .env from env_example.txt...
        copy env_example.txt .env >nul
        echo ✓ .env file created
        echo.
        echo IMPORTANT: Please edit .env file and add your API credentials!
        echo You can open it with Notepad by running:
        echo   notepad .env
    ) else (
        echo ERROR: env_example.txt not found!
        echo Please create a .env file manually.
    )
)
echo.

echo ========================================
echo   Setup Complete!
echo ========================================
echo.
echo Next steps:
echo 1. Edit .env file with your Azure OpenAI credentials
echo 2. Run launch_app.bat to start the application
echo.
echo For detailed instructions, see WINDOWS_SETUP_GUIDE.md
echo.
pause

