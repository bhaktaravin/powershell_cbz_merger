@echo off
:: Batch file to run the CBZFileMerger PowerShell script

:: Define the path to the PowerShell script
set scriptPath=C:\CBZFilesMerger\CBZFileMerger.ps1

:: Check if PowerShell is available
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo PowerShell is not installed or not in PATH. Exiting.
    exit /b 1
)

:: Run the PowerShell script
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%scriptPath%"

:: Pause to display output
pause