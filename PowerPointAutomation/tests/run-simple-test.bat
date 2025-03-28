@echo off
REM PowerPoint Simple Test Runner with Memory Optimizations
REM This batch file runs a simple test presentation using PowerShell

echo Setting up environment for PowerPoint Automation test...

REM Increase working set size for PowerShell
powershell -Command "$proc = Get-Process -Id $pid; $proc.MinWorkingSet = 204800; $proc.MaxWorkingSet = 1048576000;"

echo Environment configured for optimal memory usage.
echo Creating test presentation with PowerShell...

REM Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "%~dp0docs\simple-powerpoint-test.ps1"

echo.
echo PowerPoint test completed.
echo If you experienced memory issues, please run the IncreaseProcessMemory.ps1 script as administrator.
echo.

pause 