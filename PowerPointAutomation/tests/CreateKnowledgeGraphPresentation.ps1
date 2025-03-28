# Knowledge Graph PowerPoint Generation Script
# This script creates a complete knowledge graph presentation with all the features
# using PowerShell directly rather than requiring the C# project to be built

# Load System.Drawing assembly for color manipulation
Add-Type -AssemblyName System.Drawing

# Get the current directory and construct a fully qualified output path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$outputPath = Join-Path $projectDir "docs\output\KnowledgeGraphPresentation.pptx"

# Create output directory if it doesn't exist
$outputDir = Split-Path -Parent $outputPath
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Green
}

Write-Host "Creating Knowledge Graph presentation..." -ForegroundColor Cyan
Write-Host "Output will be saved to: $outputPath" -ForegroundColor Green

# Theme colors (for consistent branding)
$primaryColorRGB = [System.Drawing.Color]::FromArgb(31, 73, 125).ToArgb()    # Dark blue
$secondaryColorRGB = [System.Drawing.Color]::FromArgb(68, 114, 196).ToArgb() # Medium blue
$accentColorRGB = [System.Drawing.Color]::FromArgb(237, 125, 49).ToArgb()    # Orange
$lightColorRGB = [System.Drawing.Color]::FromArgb(242, 242, 242).ToArgb()    # Light gray

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
    
    # Set presentation properties
    $presentation.PageSetup.SlideSize = 4 # ppSlideSizeOnScreen16x9 = 4
    
    # === Create Title Slide ===
    $titleSlide = $presentation.Slides.Add(1, 1) # 1 = Title slide layout
    Write-Host "Title slide created" -ForegroundColor Green
    
    # Add title text
    $titleShape = $titleSlide.Shapes.Title
    $titleShape.TextFrame.TextRange.Text = "Knowledge Graphs"
    $titleShape.TextFrame.TextRange.Font.Size = 44
    $titleShape.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add subtitle text
    $subtitleShape = $titleSlide.Shapes.Item(2)
    $subtitleShape.TextFrame.TextRange.Text = "A Comprehensive Introduction"
    $subtitleShape.TextFrame.TextRange.Font.Size = 32
    $subtitleShape.TextFrame.TextRange.Font.Color.RGB = $secondaryColorRGB
    
    # Add presenter line
    $presenterShape = $titleSlide.Shapes.AddTextbox(1, 200, 400, 400, 40) # msoTextOrientationHorizontal = 1
    $presenterShape.TextFrame.TextRange.Text = "Presented by PowerPoint Automation"
    $presenterShape.TextFrame.TextRange.Font.Size = 20
    $presenterShape.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::DarkGray.ToArgb()
    $presenterShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter = 2
    
    # Add date
    $dateShape = $titleSlide.Shapes.AddTextbox(1, 200, 500, 400, 40)
    $dateShape.TextFrame.TextRange.Text = Get-Date -Format "MMMM d, yyyy"
    $dateShape.TextFrame.TextRange.Font.Size = 16
    $dateShape.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::DarkGray.ToArgb()
    $dateShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter = 2
    
    # === Create Introduction Slide ===
    $introSlide = $presentation.Slides.Add(2, 2) # 2 = Title and content layout 
    Write-Host "Introduction slide created" -ForegroundColor Green
    
    # Add title
    $introSlide.Shapes.Title.TextFrame.TextRange.Text = "Introduction to Knowledge Graphs"
    $introSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add content with bullet points
    $contentShape = $introSlide.Shapes.Item(2)
    $contentText = $contentShape.TextFrame.TextRange
    $contentText.Text = "Knowledge graphs represent information as interconnected entities and relationships"
    
    # Add bullet points
    $bulletPoints = @(
        "Semantic networks that represent real-world entities (objects, events, concepts)",
        "Bridge structured and unstructured data for human and machine interpretation",
        "Enable sophisticated reasoning, discovery, and analysis capabilities",
        "Create a flexible yet robust foundation for knowledge management"
    )
    
    foreach ($point in $bulletPoints) {
        $contentText.Paragraphs(-1).Text += [char]13 + $point
    }
    
    # Format bullet points
    for ($i = 1; $i -le $bulletPoints.Count + 1; $i++) {
        $contentText.Paragraphs($i).ParagraphFormat.Bullet.Type = 1 # ppBulletUnnumbered = 1
        $contentText.Paragraphs($i).Font.Size = 24
    }
    
    # === Create Core Components Slide ===
    $componentsSlide = $presentation.Slides.Add(3, 2)
    Write-Host "Core Components slide created" -ForegroundColor Green
    
    # Add title
    $componentsSlide.Shapes.Title.TextFrame.TextRange.Text = "Core Components of Knowledge Graphs"
    $componentsSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add content with bullet points
    $contentShape = $componentsSlide.Shapes.Item(2)
    $contentText = $contentShape.TextFrame.TextRange
    $contentText.Text = ""
    
    # Add bullet points
    $bulletPoints = @(
        "Entities (nodes): Represent real-world objects, concepts, or ideas",
        "Relationships (edges): Define how entities are connected",
        "Properties: Attributes that describe entities or relationships",
        "Ontologies: Formal specifications of conceptualization",
        "Taxonomies: Hierarchical organization of entities",
        "Inference rules: Logic for deriving new knowledge"
    )
    
    foreach ($point in $bulletPoints) {
        $contentText.Paragraphs(-1).Text += $point + [char]13
    }
    
    # Format bullet points
    for ($i = 1; $i -le $bulletPoints.Count; $i++) {
        $contentText.Paragraphs($i).ParagraphFormat.Bullet.Type = 1 # ppBulletUnnumbered = 1
        $contentText.Paragraphs($i).Font.Size = 24
    }
    
    # === Create Structural Example Slide ===
    $exampleSlide = $presentation.Slides.Add(4, 2)
    Write-Host "Structural Example slide created" -ForegroundColor Green
    
    # Add title
    $exampleSlide.Shapes.Title.TextFrame.TextRange.Text = "Knowledge Graph Structural Example"
    $exampleSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add a simple diagram using shapes and connectors
    # Create nodes (entities)
    $personNode = $exampleSlide.Shapes.AddShape(9, 100, 200, 100, 50) # msoShapeRoundedRectangle = 9
    $personNode.Fill.ForeColor.RGB = $secondaryColorRGB
    $personNode.TextFrame.TextRange.Text = "Person"
    $personNode.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::White.ToArgb()
    
    $cityNode = $exampleSlide.Shapes.AddShape(9, 400, 150, 100, 50)
    $cityNode.Fill.ForeColor.RGB = $secondaryColorRGB
    $cityNode.TextFrame.TextRange.Text = "City"
    $cityNode.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::White.ToArgb()
    
    $companyNode = $exampleSlide.Shapes.AddShape(9, 400, 250, 100, 50)
    $companyNode.Fill.ForeColor.RGB = $secondaryColorRGB
    $companyNode.TextFrame.TextRange.Text = "Company"
    $companyNode.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::White.ToArgb()
    
    # Create connectors (relationships)
    $livesInConnector = $exampleSlide.Shapes.AddConnector(2, 0, 0, 0, 0) # msoConnectorStraight = 2
    $livesInConnector.ConnectorFormat.BeginConnect($personNode, 1)
    $livesInConnector.ConnectorFormat.EndConnect($cityNode, 3)
    $livesInConnector.Line.ForeColor.RGB = $accentColorRGB
    $livesInConnector.Line.Weight = 2
    
    # Add label to connector
    $livesInLabel = $exampleSlide.Shapes.AddTextbox(1, 250, 150, 100, 30)
    $livesInLabel.TextFrame.TextRange.Text = "Lives In"
    $livesInLabel.TextFrame.TextRange.Font.Color.RGB = $accentColorRGB
    $livesInLabel.TextFrame.TextRange.Font.Bold = 1 # msoTrue = 1
    
    $worksForConnector = $exampleSlide.Shapes.AddConnector(2, 0, 0, 0, 0)
    $worksForConnector.ConnectorFormat.BeginConnect($personNode, 1)
    $worksForConnector.ConnectorFormat.EndConnect($companyNode, 3)
    $worksForConnector.Line.ForeColor.RGB = $accentColorRGB
    $worksForConnector.Line.Weight = 2
    
    # Add label to connector
    $worksForLabel = $exampleSlide.Shapes.AddTextbox(1, 250, 250, 100, 30)
    $worksForLabel.TextFrame.TextRange.Text = "Works For"
    $worksForLabel.TextFrame.TextRange.Font.Color.RGB = $accentColorRGB
    $worksForLabel.TextFrame.TextRange.Font.Bold = 1
    
    # === Create Applications Slide ===
    $applicationsSlide = $presentation.Slides.Add(5, 2)
    Write-Host "Applications slide created" -ForegroundColor Green
    
    # Add title
    $applicationsSlide.Shapes.Title.TextFrame.TextRange.Text = "Applications of Knowledge Graphs"
    $applicationsSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add content with bullet points
    $contentShape = $applicationsSlide.Shapes.Item(2)
    $contentText = $contentShape.TextFrame.TextRange
    $contentText.Text = ""
    
    # Add bullet points
    $bulletPoints = @(
        "Search engines (Google Knowledge Graph, Bing Knowledge Graph)",
        "Virtual assistants (Amazon Alexa, Apple Siri, Google Assistant)",
        "Recommendation systems (Netflix, Amazon, Spotify)",
        "Enterprise knowledge management",
        "Drug discovery and healthcare analytics",
        "Fraud detection and financial analytics"
    )
    
    foreach ($point in $bulletPoints) {
        $contentText.Paragraphs(-1).Text += $point + [char]13
    }
    
    # Format bullet points
    for ($i = 1; $i -le $bulletPoints.Count; $i++) {
        $contentText.Paragraphs($i).ParagraphFormat.Bullet.Type = 1 # ppBulletUnnumbered = 1
        $contentText.Paragraphs($i).Font.Size = 24
    }
    
    # === Create Future Directions Slide ===
    $futureSlide = $presentation.Slides.Add(6, 2)
    Write-Host "Future Directions slide created" -ForegroundColor Green
    
    # Add title
    $futureSlide.Shapes.Title.TextFrame.TextRange.Text = "Future Directions in Knowledge Graphs"
    $futureSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add content with bullet points
    $contentShape = $futureSlide.Shapes.Item(2)
    $contentText = $contentShape.TextFrame.TextRange
    $contentText.Text = ""
    
    # Add bullet points
    $bulletPoints = @(
        "Integration with large language models (LLMs)",
        "Multimodal knowledge representation (text, images, audio)",
        "Self-evolving knowledge graphs with automated updates",
        "Federated knowledge graphs across organizational boundaries",
        "Enhanced reasoning capabilities with neural-symbolic integration",
        "Quantum computing approaches to knowledge representation"
    )
    
    foreach ($point in $bulletPoints) {
        $contentText.Paragraphs(-1).Text += $point + [char]13
    }
    
    # Format bullet points
    for ($i = 1; $i -le $bulletPoints.Count; $i++) {
        $contentText.Paragraphs($i).ParagraphFormat.Bullet.Type = 1 # ppBulletUnnumbered = 1
        $contentText.Paragraphs($i).Font.Size = 24
    }
    
    # === Create Conclusion Slide ===
    $conclusionSlide = $presentation.Slides.Add(7, 1) # Use title slide layout for conclusion
    Write-Host "Conclusion slide created" -ForegroundColor Green
    
    # Add title
    $conclusionSlide.Shapes.Title.TextFrame.TextRange.Text = "Knowledge Graphs: The Future of Intelligent Data"
    $conclusionSlide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = $primaryColorRGB
    
    # Add conclusion text
    $conclusionText = $conclusionSlide.Shapes.Item(2)
    $conclusionText.TextFrame.TextRange.Text = "Connecting data, context, and meaning for the next generation of intelligent applications"
    $conclusionText.TextFrame.TextRange.Font.Size = 28
    $conclusionText.TextFrame.TextRange.Font.Color.RGB = $secondaryColorRGB
    
    # Add footer to all slides
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        if ($i -ne 1) { # Skip title slide
            $slide = $presentation.Slides.Item($i)
            
            # Add footer
            $footerShape = $slide.Shapes.AddTextbox(1, 20, 500, 400, 20)
            $footerShape.TextFrame.TextRange.Text = "Knowledge Graph Presentation | " + (Get-Date -Format "MMMM yyyy")
            $footerShape.TextFrame.TextRange.Font.Size = 10
            $footerShape.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::Gray.ToArgb()
            
            # Add slide number
            $slideNumberShape = $slide.Shapes.AddTextbox(1, 680, 500, 40, 20)
            $slideNumberShape.TextFrame.TextRange.Text = $i.ToString() + " / " + $presentation.Slides.Count
            $slideNumberShape.TextFrame.TextRange.Font.Size = 10
            $slideNumberShape.TextFrame.TextRange.Font.Color.RGB = [System.Drawing.Color]::Gray.ToArgb()
            $slideNumberShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignRight = 2
        }
    }
    
    # Add slide transitions
    foreach ($slide in $presentation.Slides) {
        # Set transition type
        $slide.SlideShowTransition.EntryEffect = 3587 # ppEffectFade = 3587
        $slide.SlideShowTransition.Duration = 1 # 1 second
        
        # Advance on click
        $slide.SlideShowTransition.AdvanceOnTime = 0 # msoFalse = 0
        $slide.SlideShowTransition.AdvanceOnClick = 1 # msoTrue = 1
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

Write-Host "Knowledge Graph Presentation created successfully." -ForegroundColor Green 