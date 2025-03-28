@echo off
REM Full Knowledge Graph PowerPoint Generator with Memory Optimizations
REM This batch file builds and runs the C# PowerPoint automation

echo Setting up environment for PowerPoint Automation...

REM Set environment variables for .NET garbage collection
set COMPLUS_gcMemoryLimit=0xFFFFFFFF
set COMPLUS_gcServer=1
set COMPLUS_gcConcurrent=0
set COMPLUS_GCLOHCompact=1

echo Environment configured for optimal memory usage.

REM Check if MSBuild exists in the standard Visual Studio 2022 location
set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
)
if not exist %MSBUILD_PATH% (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
)
if not exist %MSBUILD_PATH% (
    echo ERROR: Could not find MSBuild.exe. Please ensure Visual Studio 2022 is installed.
    pause
    exit /b 1
)

echo Building PowerPoint Automation project...
%MSBUILD_PATH% PowerPointAutomation.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo Build completed successfully.

REM Ensure output directory exists
if not exist "..\PowerPointAutomation\docs\output" mkdir "..\PowerPointAutomation\docs\output"

REM Check for slide parameter
if "%1"=="slide" (
    if "%2"=="" (
        echo Usage: run-full-presentation.bat slide [slideNumber]
        echo Available slides:
        echo   1 - Title slide
        echo   2 - Introduction slide
        echo   3 - Core Components slide
        echo   4 - Structural Example slide
        echo   5 - Applications slide
        echo   6 - Future Directions slide
        echo   7 - Conclusion slide
        pause
        exit /b 1
    )
    
    echo Running PowerPoint Automation for slide %2 only...
    PowerPointAutomation\bin\Debug\PowerPointAutomation.exe slide %2 PowerPointAutomation\docs\output\KnowledgeGraphPresentation_Slide%2.pptx
) else (
    echo Creating full Knowledge Graph presentation...
    PowerPointAutomation\bin\Debug\PowerPointAutomation.exe
)

echo.
echo PowerPoint Automation completed.
echo.

pause 