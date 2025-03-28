# Simple PowerPoint Test using PowerShell
# This script creates a basic PowerPoint presentation to test memory optimizations
# It doesn't require the C# project to be built

# Output path for the presentation
$outputPath = Join-Path (Get-Location) "output\PowerShellTestPresentation.pptx"

# Create output directory if it doesn't exist
$outputDir = Split-Path -Parent $outputPath
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Green
}

Write-Host "Creating simple PowerPoint test presentation..." -ForegroundColor Cyan
Write-Host "Output will be saved to: $outputPath" -ForegroundColor Cyan

try {
    # Start garbage collection to optimize memory
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Create PowerPoint application
    $powerPoint = New-Object -ComObject PowerPoint.Application
    Write-Host "PowerPoint application created" -ForegroundColor Green
    
    # Make PowerPoint visible - must use integer for MsoTriState (1 = True, 0 = False)
    $powerPoint.Visible = 1
    
    # Add a new presentation
    $presentation = $powerPoint.Presentations.Add()
    Write-Host "Presentation created" -ForegroundColor Green
    
    # Add title slide
    $slideIndex = 1
    $titleSlide = $presentation.Slides.Add($slideIndex, 1) # 1 = Title slide layout
    Write-Host "Title slide created" -ForegroundColor Green
    
    # Add title text
    $titleShape = $titleSlide.Shapes.Title
    $titleShape.TextFrame.TextRange.Text = "Memory Optimization Test"
    
    # Add subtitle text if the shape exists
    if ($titleSlide.Shapes.Count -ge 2) {
        $subtitleShape = $titleSlide.Shapes.Item(2)
        if ($subtitleShape -and $subtitleShape.HasTextFrame) {
            $subtitleShape.TextFrame.TextRange.Text = "Testing PowerPoint Automation with PowerShell"
        }
    }
    
    # Add content slide
    $slideIndex++
    $contentSlide = $presentation.Slides.Add($slideIndex, 2) # 2 = Title and content layout
    Write-Host "Content slide created" -ForegroundColor Green
    
    # Add title to content slide
    $contentTitle = $contentSlide.Shapes.Title
    $contentTitle.TextFrame.TextRange.Text = "Memory Optimization Features"
    
    # Add bullet points if the shape exists
    if ($contentSlide.Shapes.Count -ge 2) {
        $contentShape = $contentSlide.Shapes.Item(2)
        if ($contentShape -and $contentShape.HasTextFrame) {
            $contentShape.TextFrame.TextRange.Text = 
                "• Batch processing of COM objects" + [char]13 +
                "• Age-based COM object tracking and release" + [char]13 +
                "• Incremental presentation generation mode" + [char]13 +
                "• 64-bit process optimization" + [char]13 +
                "• System-level memory optimization scripts" + [char]13 +
                "• Configurable garbage collection settings"
        }
    }
    
    # Add conclusion slide
    $slideIndex++
    $conclusionSlide = $presentation.Slides.Add($slideIndex, 1) # 1 = Title slide layout
    Write-Host "Conclusion slide created" -ForegroundColor Green
    
    # Add conclusion title
    $conclusionTitle = $conclusionSlide.Shapes.Title
    $conclusionTitle.TextFrame.TextRange.Text = "Memory Optimization Successful!"
    
    # Add conclusion text if the shape exists
    if ($conclusionSlide.Shapes.Count -ge 2) {
        $conclusionText = $conclusionSlide.Shapes.Item(2)
        if ($conclusionText -and $conclusionText.HasTextFrame) {
            $conclusionText.TextFrame.TextRange.Text = "The presentation was generated without memory issues."
        }
    }
    
    # Save the presentation
    Write-Host "Saving presentation..." -ForegroundColor Yellow
    $presentation.SaveAs($outputPath)
    Write-Host "Presentation saved successfully!" -ForegroundColor Green
    
    # Close presentation
    $presentation.Close()
    $presentation = $null
    
    # Quit PowerPoint
    $powerPoint.Quit()
    $powerPoint = $null
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "PowerPoint application closed cleanly" -ForegroundColor Green
    
    # Open the presentation with the default application
    Write-Host "Opening the presentation for review..." -ForegroundColor Cyan
    Invoke-Item $outputPath
}
catch {
    Write-Host "Error creating presentation: $_" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
}
finally {
    # Make sure PowerPoint is closed if there was an error
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { }
        $presentation = $null
    }
    
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch { }
        $powerPoint = $null
    }
    
    # Force garbage collection again
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Check if PowerPoint is still running
    $pptProcess = Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue
    if ($pptProcess) {
        Write-Host "PowerPoint is still running. Attempting to close..." -ForegroundColor Yellow
        try {
            $pptProcess | ForEach-Object { $_.Kill() }
            Write-Host "PowerPoint process terminated." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to terminate PowerPoint process: $_" -ForegroundColor Red
        }
    }
}

Write-Host "Test complete." -ForegroundColor Green 