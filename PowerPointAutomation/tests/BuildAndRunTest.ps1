# PowerPoint Automation Build and Test Script
# This script compiles the solution and runs the simple test

Write-Host "Building PowerPoint Automation solution..." -ForegroundColor Cyan

# Locate MSBuild.exe
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2019\*\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\*\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuildPath = $null
foreach ($path in $msbuildPaths) {
    $resolvedPaths = Resolve-Path -Path $path -ErrorAction SilentlyContinue
    if ($resolvedPaths -and $resolvedPaths.Count -gt 0) {
        $msbuildPath = $resolvedPaths[0].Path
        break
    }
}

if (-not $msbuildPath) {
    # Try to use vswhere to find MSBuild
    $vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswherePath) {
        $vsPath = & $vswherePath -latest -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (-not (Test-Path $msbuildPath)) {
                $msbuildPath = Join-Path $vsPath "MSBuild\15.0\Bin\MSBuild.exe"
            }
        }
    }
}

if (-not $msbuildPath -or -not (Test-Path $msbuildPath)) {
    Write-Host "Could not find MSBuild.exe. Please ensure Visual Studio is installed." -ForegroundColor Red
    Write-Host "Trying to find any existing compiled executable..." -ForegroundColor Yellow
}
else {
    # Find the solution file
    $solutionPath = Get-ChildItem -Path . -Filter "*.sln" | Select-Object -First 1 -ExpandProperty FullName
    
    if (-not $solutionPath) {
        $solutionPath = Get-ChildItem -Path .. -Filter "*.sln" | Select-Object -First 1 -ExpandProperty FullName
    }
    
    if (-not $solutionPath) {
        Write-Host "Could not find a solution file. Please run this script from the solution directory." -ForegroundColor Red
    }
    else {
        # Build the solution
        Write-Host "Building solution: $solutionPath" -ForegroundColor Cyan
        Write-Host "Using MSBuild: $msbuildPath" -ForegroundColor Cyan
        
        # Build for x64 platform
        & $msbuildPath $solutionPath /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Build failed with exit code $LASTEXITCODE." -ForegroundColor Red
        }
        else {
            Write-Host "Build completed successfully." -ForegroundColor Green
        }
    }
}

# Find the executable
$exePath = $null
$searchPaths = @(
    ".\PowerPointAutomation\bin\Debug\PowerPointAutomation.exe",
    ".\PowerPointAutomation\bin\Release\PowerPointAutomation.exe",
    ".\PowerPointAutomation\bin\x64\Debug\PowerPointAutomation.exe",
    ".\PowerPointAutomation\bin\x64\Release\PowerPointAutomation.exe"
)

foreach ($path in $searchPaths) {
    if (Test-Path $path) {
        $exePath = $path
        break
    }
}

if (-not $exePath) {
    Write-Host "Could not find PowerPointAutomation.exe. The build may have failed or the file is in a different location." -ForegroundColor Red
    exit 1
}

# Create output directory if it doesn't exist
$outputDir = ".\PowerPointAutomation\docs\output"
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Green
}

$outputPath = Join-Path $outputDir "MemoryTestPresentation.pptx"

# Set environment variables for memory optimization
$env:COMPLUS_gcMemoryLimit = "0xFFFFFFFF"
$env:COMPLUS_gcServer = "1"
$env:COMPLUS_gcConcurrent = "0"
$env:COMPLUS_GCLOHCompact = "1"

# Run the simple test
Write-Host "Running simple test presentation..." -ForegroundColor Cyan
Write-Host "Using executable: $exePath" -ForegroundColor Cyan
Write-Host "Output will be saved to: $outputPath" -ForegroundColor Cyan

& $exePath simple $outputPath

Write-Host "Test complete." -ForegroundColor Green 