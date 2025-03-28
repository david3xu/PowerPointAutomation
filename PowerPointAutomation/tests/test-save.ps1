param(
    [int]$slideNumber = 2,
    [string]$outputPath = "C:\Users\jingu\source\repos\PowerPointAutomation\PowerPointAutomation\docs\output\Slide$slideNumber.pptx"
)

Write-Host "Starting PowerShell PowerPoint generation script..."
Write-Host "Creating slide #$slideNumber"
Write-Host "Will save to: $outputPath"

# Ensure the output directory exists
$outputDir = Split-Path -Parent $outputPath
Write-Host "Output directory: $outputDir"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force
    Write-Host "Created output directory: $outputDir"
}

try {
    # Create a new PowerPoint application instance
    $ppApp = New-Object -ComObject PowerPoint.Application
    $ppApp.Visible = 1
    
    # Create a new presentation
    $presentation = $ppApp.Presentations.Add(1)
    
    # Generate the appropriate slide based on slideNumber
    switch ($slideNumber) {
        1 {
            # Title slide
            Write-Host "Creating Title slide..."
            $slide = $presentation.Slides.Add(1, 1) # Position 1, Layout 1 = Title
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "Knowledge Graphs"
            
            # Add subtitle
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # Subtitle
                        $textFrame = $shape.TextFrame
                        $textFrame.TextRange.Text = "Understanding the Foundation of Semantic Data"
                        break
                    }
                }
            }
            
            # Add author information at the bottom
            $authorShape = $slide.Shapes.AddTextbox(1, 50, 400, 500, 50)
            $authorShape.TextFrame.TextRange.Text = "Dr. Jane Smith`nAI Research Institute"
            $authorShape.TextFrame.TextRange.Font.Size = 18
            $authorShape.TextFrame.TextRange.Font.Color.RGB = 5592405 # Dark blue
        }
        2 {
            # Introduction slide
            Write-Host "Creating Introduction slide..."
            $slide = $presentation.Slides.Add(1, 2) # Position 1, Layout 2 = Title and Content
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "What is a Knowledge Graph?"
            
            # Add content
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # ppPlaceholderBody
                        $textFrame = $shape.TextFrame
                        $textFrame.TextRange.Text = "A knowledge graph is a network of entities, their semantic types, properties, and relationships. " +
                        "It integrates data from multiple sources and enables machines to understand the semantics (meaning) " + 
                        "of data in a way that's closer to human understanding.

Knowledge graphs form the foundation of many modern AI systems, including search engines, " +
                        "virtual assistants, recommendation systems, and data integration platforms."
                        break
                    }
                }
            }
        }
        3 {
            # Core Components slide
            Write-Host "Creating Core Components slide..."
            $slide = $presentation.Slides.Add(1, 2) # Position 1, Layout 2 = Title and Content
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "Core Components"
            
            # Add bulleted list content
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # ppPlaceholderBody
                        $textFrame = $shape.TextFrame
                        $textRange = $textFrame.TextRange
                        
                        # Set up bullet points
                        $bulletPoints = @(
                            "Entities (Nodes): Real-world objects, concepts, or events represented in the graph",
                            "Relationships (Edges): Connections between entities that describe how they relate to each other",
                            "Properties: Attributes that describe entities or relationships",
                            "Ontologies: Formal definitions of types, properties, and relationships",
                            "Inference Rules: Logical rules that derive new facts from existing ones"
                        )
                        
                        # Add each bullet point
                        for ($i = 0; $i -lt $bulletPoints.Count; $i++) {
                            if ($i -gt 0) {
                                # Add a paragraph for each subsequent bullet point
                                $paraIndex = $textRange.Paragraphs().Count
                                $newPara = $textRange.Paragraphs($paraIndex).InsertAfter("`r`n")
                                $bulletPara = $textRange.Paragraphs($paraIndex + 1)
                                $bulletPara.Text = $bulletPoints[$i]
                            } else {
                                # First bullet point
                                $textRange.Text = $bulletPoints[$i]
                            }
                        }
                        break
                    }
                }
            }
        }
        4 {
            # Structural Example slide (Diagram)
            Write-Host "Creating Structural Example slide (diagram)..."
            $slide = $presentation.Slides.Add(1, 11) # Position 1, Layout 11 = Blank
            
            # Add title manually since it's a blank slide
            $titleShape = $slide.Shapes.AddTextbox(1, 50, 50, 600, 50)
            $titleShape.TextFrame.TextRange.Text = "Structural Example"
            $titleShape.TextFrame.TextRange.Font.Size = 32
            $titleShape.TextFrame.TextRange.Font.Bold = 1
            
            # Add subtitle
            $subtitleShape = $slide.Shapes.AddTextbox(1, 50, 100, 600, 30)
            $subtitleShape.TextFrame.TextRange.Text = "A simple knowledge graph showing entities and relationships"
            $subtitleShape.TextFrame.TextRange.Font.Size = 20
            $subtitleShape.TextFrame.TextRange.Font.Italic = 1
            
            # Create entities (nodes)
            $entities = @{
                "Person" = [PSCustomObject]@{X = 150; Y = 200; Width = 100; Height = 50; Color = 13395456} # Blue
                "Movie" = [PSCustomObject]@{X = 400; Y = 200; Width = 100; Height = 50; Color = 5287936} # Green
                "Actor" = [PSCustomObject]@{X = 150; Y = 350; Width = 100; Height = 50; Color = 13395456} # Blue
                "Director" = [PSCustomObject]@{X = 400; Y = 350; Width = 100; Height = 50; Color = 13395456} # Blue
            }
            
            # Create the entity shapes - using msoShapeRectangle (1)
            $entityShapes = @{}
            foreach ($entity in $entities.Keys) {
                $shape = $slide.Shapes.AddShape(1, $entities[$entity].X, $entities[$entity].Y, 
                                               $entities[$entity].Width, $entities[$entity].Height)
                $shape.Fill.ForeColor.RGB = $entities[$entity].Color
                $shape.Line.ForeColor.RGB = 0 # Black
                $shape.TextFrame.TextRange.Text = $entity
                $shape.TextFrame.TextRange.Font.Size = 14
                $shape.TextFrame.TextRange.Font.Bold = 1
                # Use different methods for anchoring
                $shape.TextFrame.WordWrap = 1
                $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Center
                $entityShapes[$entity] = $shape
            }
            
            # Add relationship lines with labels - using msoConnectorStraight (1) instead of 2
            # Person -> Movie (watches)
            $connector1 = $slide.Shapes.AddConnector(1, $entityShapes["Person"].Left + $entityShapes["Person"].Width, 
                                                  $entityShapes["Person"].Top + 25, 
                                                  $entityShapes["Movie"].Left, 
                                                  $entityShapes["Movie"].Top + 25)
            $connector1.Line.ForeColor.RGB = 0 # Black
            $connector1.Line.Weight = 1.5
            
            # Add relationship label
            $label1 = $slide.Shapes.AddTextbox(1, 275, 175, 100, 30)
            $label1.TextFrame.TextRange.Text = "watches"
            $label1.TextFrame.TextRange.Font.Size = 12
            $label1.TextFrame.TextRange.Font.Italic = 1
            
            # Actor -> Movie (acts_in)
            $connector2 = $slide.Shapes.AddConnector(1, $entityShapes["Actor"].Left + $entityShapes["Actor"].Width, 
                                                  $entityShapes["Actor"].Top + 25, 
                                                  $entityShapes["Movie"].Left, 
                                                  $entityShapes["Movie"].Top + 40)
            $connector2.Line.ForeColor.RGB = 0 # Black
            $connector2.Line.Weight = 1.5
            
            # Add relationship label
            $label2 = $slide.Shapes.AddTextbox(1, 275, 325, 100, 30)
            $label2.TextFrame.TextRange.Text = "acts_in"
            $label2.TextFrame.TextRange.Font.Size = 12
            $label2.TextFrame.TextRange.Font.Italic = 1
            
            # Director -> Movie (directs)
            $connector3 = $slide.Shapes.AddConnector(1, $entityShapes["Director"].Left + $entityShapes["Director"].Width/2, 
                                                  $entityShapes["Director"].Top, 
                                                  $entityShapes["Movie"].Left + $entityShapes["Movie"].Width/2, 
                                                  $entityShapes["Movie"].Top + $entityShapes["Movie"].Height)
            $connector3.Line.ForeColor.RGB = 0 # Black
            $connector3.Line.Weight = 1.5
            
            # Add relationship label
            $label3 = $slide.Shapes.AddTextbox(1, 420, 275, 100, 30)
            $label3.TextFrame.TextRange.Text = "directs"
            $label3.TextFrame.TextRange.Font.Size = 12
            $label3.TextFrame.TextRange.Font.Italic = 1
            
            # Add a legend
            $legendBox = $slide.Shapes.AddShape(1, 520, 150, 150, 100)
            $legendBox.Fill.ForeColor.RGB = 16777215 # White
            $legendBox.Line.ForeColor.RGB = 0 # Black
            
            $legendTitle = $slide.Shapes.AddTextbox(1, 530, 160, 130, 20)
            $legendTitle.TextFrame.TextRange.Text = "Legend"
            $legendTitle.TextFrame.TextRange.Font.Bold = 1
            
            # Entity color samples
            $entitySample1 = $slide.Shapes.AddShape(1, 530, 190, 20, 20)
            $entitySample1.Fill.ForeColor.RGB = 13395456 # Blue
            
            $entityLabel1 = $slide.Shapes.AddTextbox(1, 560, 190, 100, 20)
            $entityLabel1.TextFrame.TextRange.Text = "Person Type"
            
            $entitySample2 = $slide.Shapes.AddShape(1, 530, 220, 20, 20)
            $entitySample2.Fill.ForeColor.RGB = 5287936 # Green
            
            $entityLabel2 = $slide.Shapes.AddTextbox(1, 560, 220, 100, 20)
            $entityLabel2.TextFrame.TextRange.Text = "Content Type"
        }
        5 {
            # Applications slide
            Write-Host "Creating Applications slide..."
            $slide = $presentation.Slides.Add(1, 2) # Position 1, Layout 2 = Title and Content
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "Applications of Knowledge Graphs"
            
            # Add bulleted list content
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # ppPlaceholderBody
                        $textFrame = $shape.TextFrame
                        $textRange = $textFrame.TextRange
                        
                        # Set up bullet points
                        $bulletPoints = @(
                            "Semantic Search: Understanding the intent and contextual meaning behind queries",
                            "Question Answering: Providing direct answers from structured knowledge",
                            "Recommendation Systems: Suggesting products, content, or connections based on relationships",
                            "Data Integration: Combining heterogeneous data sources with a unified semantic model",
                            "Explainable AI: Adding interpretability to AI systems through knowledge representation"
                        )
                        
                        # Add each bullet point
                        for ($i = 0; $i -lt $bulletPoints.Count; $i++) {
                            if ($i -gt 0) {
                                # Add a paragraph for each subsequent bullet point
                                $paraIndex = $textRange.Paragraphs().Count
                                $newPara = $textRange.Paragraphs($paraIndex).InsertAfter("`r`n")
                                $bulletPara = $textRange.Paragraphs($paraIndex + 1)
                                $bulletPara.Text = $bulletPoints[$i]
                            } else {
                                # First bullet point
                                $textRange.Text = $bulletPoints[$i]
                            }
                        }
                        break
                    }
                }
            }
        }
        6 {
            # Future Directions slide (two columns)
            Write-Host "Creating Future Directions slide..."
            $slide = $presentation.Slides.Add(1, 3) # Position 1, Layout 3 = Two Content
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "Future Directions"
            
            # Find the two content placeholders
            $leftContent = $null
            $rightContent = $null
            
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # Content placeholder
                        if ($shape.Left -lt 300) {
                            $leftContent = $shape
                        } else {
                            $rightContent = $shape
                        }
                    }
                }
            }
            
            if ($leftContent -and $rightContent) {
                # Left column - technologies
                $leftTextFrame = $leftContent.TextFrame
                $leftTextRange = $leftTextFrame.TextRange
                
                $technologies = @(
                    "Multimodal Knowledge Graphs",
                    "Temporal Knowledge Representation",
                    "Federated Knowledge Graphs",
                    "Neural-Symbolic Integration",
                    "Quantum Knowledge Representation"
                )
                
                # Add each bullet point to left column
                for ($i = 0; $i -lt $technologies.Count; $i++) {
                    if ($i -gt 0) {
                        # Add a paragraph for each subsequent bullet point
                        $paraIndex = $leftTextRange.Paragraphs().Count
                        $newPara = $leftTextRange.Paragraphs($paraIndex).InsertAfter("`r`n")
                        $bulletPara = $leftTextRange.Paragraphs($paraIndex + 1)
                        $bulletPara.Text = $technologies[$i]
                    } else {
                        # First bullet point
                        $leftTextRange.Text = $technologies[$i]
                    }
                }
                
                # Right column - descriptions
                $rightTextFrame = $rightContent.TextFrame
                $rightTextRange = $rightTextFrame.TextRange
                
                $descriptions = @(
                    "Incorporating images, video, audio into knowledge structures",
                    "Representing time-dependent facts and evolving relationships",
                    "Distributed knowledge graphs across organizations",
                    "Combining neural networks with symbolic reasoning",
                    "Quantum computing approaches to knowledge representation"
                )
                
                # Add each bullet point to right column
                for ($i = 0; $i -lt $descriptions.Count; $i++) {
                    if ($i -gt 0) {
                        # Add a paragraph for each subsequent bullet point
                        $paraIndex = $rightTextRange.Paragraphs().Count
                        $newPara = $rightTextRange.Paragraphs($paraIndex).InsertAfter("`r`n")
                        $bulletPara = $rightTextRange.Paragraphs($paraIndex + 1)
                        $bulletPara.Text = $descriptions[$i]
                    } else {
                        # First bullet point
                        $rightTextRange.Text = $descriptions[$i]
                    }
                }
            } else {
                Write-Host "Warning: Could not find both content placeholders for the two-column layout"
            }
        }
        7 {
            # Conclusion slide
            Write-Host "Creating Conclusion slide..."
            $slide = $presentation.Slides.Add(1, 2) # Position 1, Layout 2 = Title and Content
            
            # Add title
            $title = $slide.Shapes.Title
            $title.TextFrame.TextRange.Text = "Conclusion"
            
            # Add content
            $shapes = $slide.Shapes
            foreach ($shape in $shapes) {
                if ($shape.Type -eq 14) { # msoPlaceholder
                    if ($shape.PlaceholderFormat.Type -eq 2) { # ppPlaceholderBody
                        $textFrame = $shape.TextFrame
                        $textFrame.TextRange.Text = "Knowledge graphs provide a powerful framework for representing and connecting information in a way that both humans and machines can understand and reason with.

As AI systems continue to advance, knowledge graphs will play an increasingly important role in providing the structured knowledge foundation that enables more sophisticated understanding, reasoning, and explainability."
                        break
                    }
                }
            }
        }
        default {
            Write-Host "Invalid slide number. Please specify a number between 1 and 7."
            return
        }
    }
    
    # Save the presentation - convert path to absolute
    $absolutePath = [System.IO.Path]::GetFullPath($outputPath)
    Write-Host "Saving presentation to absolute path: $absolutePath"
    $presentation.SaveAs($absolutePath)
    Write-Host "Saved presentation successfully to: $absolutePath"
    
    # Close PowerPoint
    $presentation.Close()
    $ppApp.Quit()
    
    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
    
    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Script completed successfully!"
    
    # Try to open the presentation for verification
    if (Test-Path $absolutePath) {
        Write-Host "Opening the saved presentation..."
        Invoke-Item $absolutePath
    }
}
catch {
    Write-Host "Error: $($_)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)"
    Write-Host "Exception: $($_.Exception)"
}
