@echo off
REM PowerPoint Automation Runner with Optimal Memory Settings
REM This batch file configures and runs the PowerPoint automation with settings to avoid memory issues

echo Setting up environment for PowerPoint Automation...

REM Set higher memory limit for .NET processes
set COMPLUS_gcMemoryLimit=0xFFFFFFFF

REM Set server GC mode
set COMPLUS_gcServer=1

REM Set concurrent GC mode off for more predictable cleanup
set COMPLUS_gcConcurrent=0

REM Set large object heap compaction mode
set COMPLUS_GCLOHCompact=1

REM Increase the working set size
powershell -Command "$proc = Get-Process -Id $pid; $proc.MinWorkingSet = 204800; $proc.MaxWorkingSet = 1048576000;"

REM Create output directory if it doesn't exist
if not exist "docs\output" mkdir "docs\output"

echo Environment configured for optimal memory usage.
echo Starting PowerPoint Automation...

REM Find the correct path to the executable by checking common locations
if exist "bin\Debug\PowerPointAutomation.exe" (
    set EXEPATH=bin\Debug\PowerPointAutomation.exe
) else if exist "bin\Release\PowerPointAutomation.exe" (
    set EXEPATH=bin\Release\PowerPointAutomation.exe
) else if exist "bin\x64\Debug\PowerPointAutomation.exe" (
    set EXEPATH=bin\x64\Debug\PowerPointAutomation.exe
) else if exist "bin\x64\Release\PowerPointAutomation.exe" (
    set EXEPATH=bin\x64\Release\PowerPointAutomation.exe
) else (
    echo ERROR: Could not find PowerPointAutomation.exe
    echo Checked in bin\Debug, bin\Release, bin\x64\Debug, and bin\x64\Release
    echo Please build the solution first or specify the correct path.
    goto :end
)

echo Found executable at %EXEPATH%

REM Run the application with incremental processing option and custom output path
%EXEPATH% incremental "docs\output\KnowledgeGraphPresentation.pptx"

REM If you want to run in standard mode, use:
REM PowerPointAutomation.exe

echo.
echo PowerPoint Automation completed.
echo If you experienced memory issues, please run the IncreaseProcessMemory.ps1 script as administrator.
echo.

:end
pause 