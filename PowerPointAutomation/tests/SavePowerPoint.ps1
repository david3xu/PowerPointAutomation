
                        param(
                            [string]$outputPath
                        )
                        
                        $ppApp = New-Object -ComObject PowerPoint.Application
                        $ppApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                        
                        # Create a new presentation
                        $presentation = $ppApp.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)
                        
                        # Create a slide
                        $slide = $presentation.Slides.Add(1, 1) # Add slide at position 1 with layout 1
                        
                        # Add title
                        $title = $slide.Shapes.Title
                        $title.TextFrame.TextRange.Text = "Knowledge Graph - Test Slide"
                        
                        # Save the presentation to the specified path
                        $presentation.SaveAs($outputPath)
                        Write-Host "Saved presentation to $outputPath"
                        
                        # Close the presentation and quit PowerPoint
                        $presentation.Close()
                        $ppApp.Quit()
                        
                        # Release COM objects
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide) | Out-Null
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                        