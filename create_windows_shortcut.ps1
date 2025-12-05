# PowerShell script to create a Windows desktop shortcut
# Run this script: powershell -ExecutionPolicy Bypass -File create_windows_shortcut.ps1

$ErrorActionPreference = "Stop"

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Topic Generator.lnk"

# Paths
$batFile = Join-Path $scriptDir "launch_app.bat"
$iconPath = Join-Path $scriptDir "assets\app_icon.ico"

Write-Host "Creating Windows desktop shortcut..." -ForegroundColor Cyan
Write-Host ""

# Check if batch file exists
if (-not (Test-Path $batFile)) {
    Write-Host "ERROR: launch_app.bat not found!" -ForegroundColor Red
    Write-Host "Location: $batFile" -ForegroundColor Yellow
    exit 1
}

# Create WScript Shell object
$WScriptShell = New-Object -ComObject WScript.Shell

# Create shortcut
$shortcut = $WScriptShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $batFile
$shortcut.WorkingDirectory = $scriptDir
$shortcut.Description = "Topic Generator Web App - AI-powered Excel topic generation"
$shortcut.WindowStyle = 1  # Normal window

# Try to set icon (if available)
if (Test-Path $iconPath) {
    $shortcut.IconLocation = $iconPath
} else {
    # Use Python icon as fallback (if Python is installed)
    $pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($pythonExe) {
        $shortcut.IconLocation = "$pythonExe,0"
    }
}

$shortcut.Save()

Write-Host "Shortcut created successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Location: $shortcutPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "To use:" -ForegroundColor Yellow
Write-Host "  1. Double-click 'Topic Generator' on your Desktop" -ForegroundColor White
Write-Host "  2. A command window will open and start the web server" -ForegroundColor White
Write-Host "  3. Your browser will open automatically" -ForegroundColor White
Write-Host ""
Write-Host "To remove: Right-click the shortcut and select Delete" -ForegroundColor Gray
