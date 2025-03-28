# PowerPoint Knowledge Graph Presentation Generator
# This script generates all 7 slides of the Knowledge Graph presentation

$outputDir = "C:\Users\jingu\source\repos\PowerPointAutomation\PowerPointAutomation\docs\output"

# Ensure the output directory exists
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force
    Write-Host "Created output directory: $outputDir"
}

Write-Host "Starting Knowledge Graph Presentation Generation..."

# Generate all slides
for ($slideNumber = 1; $slideNumber -le 7; $slideNumber++) {
    $slideName = switch ($slideNumber) {
        1 { "TitleSlide" }
        2 { "IntroductionSlide" }
        3 { "CoreComponentsSlide" }
        4 { "DiagramSlide" }
        5 { "ApplicationsSlide" }
        6 { "FutureDirectionsSlide" }
        7 { "ConclusionSlide" }
    }
    
    $outputPath = "$outputDir\$slideName.pptx"
    
    Write-Host "`n============================"
    Write-Host "Generating slide #$slideNumber ($slideName)"
    Write-Host "============================"
    
    # Call the slide generation script with the current slide number and output path
    & .\test-save.ps1 -slideNumber $slideNumber -outputPath $outputPath
    
    # Wait a moment between slide generations
    Start-Sleep -Seconds 2
}

Write-Host "`nAll slides have been generated successfully!"
Write-Host "Slides are available in: $outputDir"

# Optionally open the output directory
Invoke-Item $outputDir 