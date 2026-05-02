# Build_EXE.ps1
# Run this script once to create Trade_Diary.exe
# Right-click this file and choose "Run with PowerShell"

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "   Trade Diary — EXE Builder" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
try {
    $pyver = python --version 2>&1
    Write-Host "  [OK] $pyver found" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Python not found. Install from python.org" -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
}

Write-Host "  [1/3] Installing PyInstaller and dependencies..." -ForegroundColor Yellow
pip install pyinstaller pyqt6 pandas openpyxl --quiet

Write-Host "  [2/3] Building Trade_Diary.exe ..." -ForegroundColor Yellow
Write-Host "        (This may take 1-2 minutes)" -ForegroundColor Gray

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

pyinstaller `
    --onefile `
    --windowed `
    --name "Trade_Diary" `
    --add-data "Trade_Diary_Backtest_1.xlsx;." `
    trade_diary.py

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "  [3/3] SUCCESS!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Your app is ready at:" -ForegroundColor White
    Write-Host "  $scriptDir\dist\Trade_Diary.exe" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  IMPORTANT: Copy Trade_Diary.exe AND Trade_Diary_Backtest_1.xlsx" -ForegroundColor Yellow
    Write-Host "  into the SAME folder before running." -ForegroundColor Yellow
    Write-Host ""
    
    # Open the dist folder
    Start-Process explorer.exe "$scriptDir\dist"
} else {
    Write-Host ""
    Write-Host "  [ERROR] Build failed. See errors above." -ForegroundColor Red
}

Read-Host "  Press Enter to close"
