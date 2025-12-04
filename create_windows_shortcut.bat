@echo off
REM Simple batch file to create Windows shortcut
REM This calls the PowerShell script with proper execution policy

echo Creating Windows desktop shortcut...
echo.

REM Check if PowerShell is available
powershell -Command "Get-Host" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell is not available!
    echo Please install PowerShell or create the shortcut manually.
    pause
    exit /b 1
)

REM Run PowerShell script with execution policy bypass
powershell -ExecutionPolicy Bypass -File "%~dp0create_windows_shortcut.ps1"

if errorlevel 1 (
    echo.
    echo An error occurred. Please check the messages above.
    pause
)

