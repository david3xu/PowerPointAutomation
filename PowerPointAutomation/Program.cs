using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Runtime.InteropServices;
using PowerPointAutomation.Utilities;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set process memory settings
            ProcessMemoryManager.SetupProcessMemory();
            
            // Check command line arguments
            string outputPath = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop), "KnowledgeGraphPresentation.pptx");
            
            int? slideIndex = null;
                
            // Check if a custom output path was provided
            if (args.Length > 1)
            {
                // Second argument is the output path
                outputPath = Path.GetFullPath(args[1]);
                
                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    try
                    {
                        Directory.CreateDirectory(outputDir);
                        Console.WriteLine($"Created output directory: {outputDir}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to create output directory: {ex.Message}");
                        // Fall back to desktop
                        outputPath = Path.Combine(Environment.GetFolderPath(
                            Environment.SpecialFolder.Desktop), "KnowledgeGraphPresentation.pptx");
                    }
                }
            }
            
            // Check for operation mode in first argument
            if (args.Length > 0)
            {
                if (args[0].ToLower() == "test")
                {
                    // RunCompatibilityTests();
                    return;
                }
                else if (args[0].ToLower() == "incremental")
                {
                    IncrementalPresentationGenerator.RunIncrementalPresentation(outputPath);
                    return;
                }
                else if (args[0].ToLower() == "simple")
                {
                    // RunSimpleTest(outputPath);
                    return;
                }
                else if (args[0].ToLower() == "slide" && args.Length > 2)
                {
                    // Format: PowerPointAutomation.exe slide [index] [outputPath]
                    if (int.TryParse(args[1], out int index) && index >= 1 && index <= 7)
                    {
                        slideIndex = index;
                        Console.WriteLine($"Generating only slide #{slideIndex}");
                        
                        // Get the correct output path from command line
                        if (args.Length > 2)
                        {
                            // Use the full path specified in args[2]
                            outputPath = Path.GetFullPath(args[2]);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Invalid slide index. Please specify a number between 1 and 7.");
                        Console.WriteLine("Usage: PowerPointAutomation.exe slide [index] [outputPath]");
                        return;
                    }
                }
            }
            
            Console.WriteLine($"Output will be saved to: {outputPath}");

            Console.WriteLine("Creating Knowledge Graph presentation...");

            try
            {
                if (slideIndex.HasValue)
                {
                    GenerateSingleSlide(slideIndex.Value, outputPath);
                }
                else
                {
                    // Create presentation generator instance
                    var presentationGenerator = new KnowledgeGraphPresentation();
                    presentationGenerator.OutputPath = outputPath;

                    // Ensure we have maximum memory available
                    GC.Collect(2, GCCollectionMode.Forced, true, true);
                    GC.WaitForPendingFinalizers();
                    
                    // Set timeout for presentation generation (5 minutes)
                    bool presentationCompleted = false;
                    Exception generateException = null;
                    
                    // Create a task to monitor the presentation generation
                    var generationTask = Task.Run(() => 
                    {
                        try
                        {
                            // Generate the presentation (this launches PowerPoint)
                            presentationGenerator.Generate();
                            presentationCompleted = true;
                        }
                        catch (Exception ex)
                        {
                            // Capture the exception for reporting outside the task
                            generateException = ex;
                        }
                    });
                    
                    // Wait for the task to complete with a timeout
                    bool completed = generationTask.Wait(TimeSpan.FromMinutes(5));
                    
                    // Cleanup regardless of how the generation went
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    
                    // Check the result
                    if (!completed)
                    {
                        Console.WriteLine("ERROR: Presentation generation timed out after 5 minutes.");
                        
                        // Try to perform emergency cleanup of PowerPoint processes
                        try
                        {
                            Console.WriteLine("Performing emergency cleanup of PowerPoint processes...");
                            PowerPointOperations.TerminatePowerPointProcesses();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error during emergency cleanup: {ex.Message}");
                        }
                    }
                    else if (generateException != null)
                    {
                        // Report the exception that occurred during generation
                        Console.WriteLine($"Error creating presentation: {generateException.Message}");
                        Console.WriteLine($"Stack trace: {generateException.StackTrace}");
                        
                        if (generateException.InnerException != null)
                        {
                            Console.WriteLine($"Inner exception: {generateException.InnerException.Message}");
                        }
                    }
                    else if (presentationCompleted)
                    {
                        Console.WriteLine($"Presentation successfully created at: {outputPath}");
                        
                        // Force garbage collection before attempting to open the file
                        GC.Collect(2, GCCollectionMode.Forced, true, true);
                        GC.WaitForPendingFinalizers();
                        
                        // Wait a moment to ensure file is fully saved and accessible
                        Thread.Sleep(1000);
                        
                        // Open the presentation (optional)
                        try
                        {
                            Console.WriteLine("Opening the presentation for review...");
                            System.Diagnostics.Process.Start(outputPath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not open the presentation: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical error in presentation generation process: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                
                // Ensure PowerPoint processes are terminated in case of critical failure
                try
                {
                    PowerPointOperations.TerminatePowerPointProcesses();
                }
                catch
                {
                    // Ignore errors during emergency cleanup
                }
            }
            finally
            {
                // Simple cleanup for Main method - don't reference presentation or pptApp here
                // as they're defined in GenerateSingleSlide, not here
                Console.WriteLine("Performing final cleanup with garbage collection");
                
                // Force garbage collection
                GC.Collect(2, GCCollectionMode.Forced, true, true);
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Verify PowerPoint processes are terminated
                PowerPointOperations.TerminatePowerPointProcesses();
            }
            
            // Wait for user input before closing
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Generates a presentation with only a specific slide
        /// </summary>
        /// <param name="slideIndex">The index of the slide to generate (1-7)</param>
        /// <param name="outputPath">The path where the presentation will be saved</param>
        private static void GenerateSingleSlide(int slideIndex, string outputPath)
        {
            Console.WriteLine($"Generating presentation with only slide #{slideIndex}...");
            
            Application pptApp = null;
            Presentation presentation = null;
            Slide slide = null;
            
            try
            {
                // Create PowerPoint application with strong references
                pptApp = new Application();
                // Add extra reference count to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(pptApp));
                ComReleaser.TrackObject(pptApp);
                pptApp.Visible = MsoTriState.msoTrue;
                
                // Keep application alive during operations
                GC.KeepAlive(pptApp);
                
                // Add a brief delay to ensure COM objects are stable
                Thread.Sleep(100);
                
                // Create a new presentation
                presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);
                // Add extra reference count to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(presentation));
                ComReleaser.TrackObject(presentation);
                
                // Keep presentation reference alive 
                GC.KeepAlive(presentation);
                GC.KeepAlive(pptApp);
                
                // Add a brief delay to ensure COM objects are stable
                Thread.Sleep(100);
                
                // Setup presentation
                presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen;
                
                // Keep strong references during the entire process
                GC.KeepAlive(presentation);
                GC.KeepAlive(pptApp);
                
                // Get necessary layouts
                int layoutCount = 0;
                try
                {
                    layoutCount = presentation.SlideMaster.CustomLayouts.Count;
                    Console.WriteLine($"Available layouts: {layoutCount}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading layouts: {ex.Message}");
                    layoutCount = 11; // Default layout count for PowerPoint
                }
                
                // Keep strong references
                GC.KeepAlive(presentation);
                GC.KeepAlive(pptApp);
                
                // Get layouts safely with extra reference counting
                CustomLayout titleLayout = null;
                CustomLayout contentLayout = null;
                CustomLayout blankLayout = null;
                CustomLayout twoContentLayout = null;
                
                try
                {
                    titleLayout = presentation.SlideMaster.CustomLayouts[Math.Min(1, layoutCount)]; // Title layout
                    ComReleaser.TrackObject(titleLayout);
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(titleLayout));
                    
                    // Keep references alive
                    GC.KeepAlive(titleLayout);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                    Thread.Sleep(50);
                    
                    contentLayout = presentation.SlideMaster.CustomLayouts[Math.Min(2, layoutCount)]; // Content layout
                    ComReleaser.TrackObject(contentLayout);
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(contentLayout));
                    
                    // Keep references alive
                    GC.KeepAlive(contentLayout);
                    GC.KeepAlive(titleLayout);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                    Thread.Sleep(50);
                    
                    blankLayout = presentation.SlideMaster.CustomLayouts[Math.Min(layoutCount, layoutCount > 7 ? 11 : layoutCount)]; // Blank layout or last available
                    ComReleaser.TrackObject(blankLayout);
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(blankLayout));
                    
                    // Keep references alive
                    GC.KeepAlive(blankLayout);
                    GC.KeepAlive(contentLayout);
                    GC.KeepAlive(titleLayout);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                    Thread.Sleep(50);
                    
                    twoContentLayout = presentation.SlideMaster.CustomLayouts[Math.Min(3, layoutCount)]; // Two content layout
                    ComReleaser.TrackObject(twoContentLayout);
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(twoContentLayout));
                    
                    // Keep references alive
                    GC.KeepAlive(twoContentLayout);
                    GC.KeepAlive(blankLayout);
                    GC.KeepAlive(contentLayout);
                    GC.KeepAlive(titleLayout);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error getting layouts: {ex.Message}");
                    // Continue with null layouts - slide generators will create their own shapes
                }
                
                // Disable automatic COM object cleanup during slide generation
                ComReleaser.PauseRelease();
                ComReleaser.ReleaseOldestObjects(0); // Reset any pending releases
                
                // Create the requested slide with increased exception handling
                try
                {
                    Console.WriteLine($"Creating slide {slideIndex}...");
                    
                    switch (slideIndex)
                    {
                        case 1: // Title slide
                            Console.WriteLine("Creating title slide...");
                            var titleSlideGenerator = new Slides.TitleSlide(presentation, titleLayout);
                            slide = titleSlideGenerator.Generate(
                                "Knowledge Graphs",
                                "Understanding the Foundation of Semantic Data",
                                "Dr. Jane Smith",
                                "AI Research Institute");
                            break;
                            
                        case 2: // Introduction slide
                            Console.WriteLine("Creating introduction slide...");
                            var introSlideGenerator = new Slides.IntroductionSlide(presentation, contentLayout);
                            slide = introSlideGenerator.Generate(
                                "What is a Knowledge Graph?",
                                "A knowledge graph is a network of entities, their semantic types, properties, and relationships. " +
                                "It integrates data from multiple sources and enables machines to understand the semantics (meaning) " +
                                "of data in a way that's closer to human understanding.\n\n" +
                                "Knowledge graphs form the foundation of many modern AI systems, including search engines, " +
                                "virtual assistants, recommendation systems, and data integration platforms.");
                            break;
                            
                        case 3: // Core components slide
                            Console.WriteLine("Creating core components slide...");
                            var featureSlideGenerator = new Slides.CoreFeatureSlide(presentation, contentLayout);
                            slide = featureSlideGenerator.Generate(
                                "Core Components",
                                new string[] {
                                    "Entities (Nodes): Real-world objects, concepts, or events represented in the graph",
                                    "Relationships (Edges): Connections between entities that describe how they relate to each other",
                                    "Properties: Attributes that describe entities or relationships",
                                    "Ontologies: Formal definitions of types, properties, and relationships",
                                    "Inference Rules: Logical rules that derive new facts from existing ones"
                                });
                            break;
                            
                        case 4: // Structural example slide
                            Console.WriteLine("Creating structural example slide...");
                            var diagramSlideGenerator = new Slides.DiagramSlide(presentation, blankLayout);
                            slide = diagramSlideGenerator.GenerateKnowledgeGraphDiagram(
                                "Structural Example",
                                "A simple knowledge graph showing entities and relationships");
                            break;
                            
                        case 5: // Applications slide
                            Console.WriteLine("Creating applications slide...");
                            var listSlideGenerator = new Slides.ListSlide(presentation, contentLayout);
                            slide = listSlideGenerator.GenerateBulletedList(
                                "Applications of Knowledge Graphs",
                                new string[] {
                                    "Semantic Search: Understanding the intent and contextual meaning behind queries",
                                    "Question Answering: Providing direct answers from structured knowledge",
                                    "Recommendation Systems: Suggesting products, content, or connections based on relationships",
                                    "Data Integration: Combining heterogeneous data sources with a unified semantic model",
                                    "Explainable AI: Adding interpretability to AI systems through knowledge representation"
                                });
                            break;
                            
                        case 6: // Future directions slide
                            Console.WriteLine("Creating future directions slide...");
                            var comparisonSlideGenerator = new Slides.ComparisonSlide(presentation, twoContentLayout);
                            slide = comparisonSlideGenerator.Generate(
                                "Future Directions",
                                new string[] {
                                    "Multimodal Knowledge Graphs",
                                    "Temporal Knowledge Representation",
                                    "Federated Knowledge Graphs",
                                    "Neural-Symbolic Integration",
                                    "Quantum Knowledge Representation"
                                },
                                new string[] {
                                    "Incorporating images, video, audio into knowledge structures",
                                    "Representing time-dependent facts and evolving relationships",
                                    "Distributed knowledge graphs across organizations",
                                    "Combining neural networks with symbolic reasoning",
                                    "Quantum computing approaches to knowledge representation"
                                },
                                "Current Research",
                                "Practical Applications");
                            break;
                            
                        case 7: // Conclusion slide
                            Console.WriteLine("Creating conclusion slide...");
                            var summarySlideGenerator = new Slides.SummarySlide(presentation, contentLayout);
                            slide = summarySlideGenerator.Generate(
                                "Conclusion",
                                "Knowledge graphs provide a powerful framework for representing and connecting information in a way that both humans and machines can understand and reason with.\n\n" +
                                "As AI systems continue to advance, knowledge graphs will play an increasingly important role in providing the structured knowledge foundation that enables more sophisticated understanding, reasoning, and explainability.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating slide: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                }
                
                // Resume normal COM object release
                ComReleaser.ResumeRelease();
                
                // IMPORTANT: Instead of suspending GC, we'll use strong references and careful timing
                Console.WriteLine("Maintaining strong COM references for all slide objects...");
                
                // Immediately protect the slide from GC
                if (slide != null)
                {
                    // Important: Track the slide COM object explicitly and keep a strong reference
                    ComReleaser.TrackObject(slide);
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(slide));
                    Console.WriteLine($"Slide #{slideIndex} created successfully!");
                    
                    try
                    {
                        // CRITICAL: Ensure shape collection access is wrapped in its own try/catch
                        try
                        {
                            var shapes = slide.Shapes;
                            ComReleaser.TrackObject(shapes);
                            Console.WriteLine($"Successfully accessed slide.Shapes collection. Count: {shapes.Count}");
                            
                            int shapeCount = 0;
                            List<PowerPointShape> shapeList = new List<PowerPointShape>();
                            
                            // Limit the number of shapes we process to avoid overload
                            int maxShapesToProcess = Math.Min(shapes.Count, 50);
                            Console.WriteLine($"Processing up to {maxShapesToProcess} shapes");
                            
                            for (int i = 1; i <= maxShapesToProcess; i++)
                            {
                                try 
                                {
                                    PowerPointShape shape = shapes[i];
                                    if (shape != null)
                                    {
                                        shapeCount++;
                                        if (shapeCount % 5 == 0)
                                        {
                                            Console.WriteLine($"  Processed {shapeCount} shapes so far");
                                            
                                            // Force references to stay alive periodically
                                            GC.KeepAlive(shape);
                                            GC.KeepAlive(shapes);
                                            GC.KeepAlive(slide);
                                            GC.KeepAlive(presentation);
                                            GC.KeepAlive(pptApp);
                                            
                                            // Give the COM objects a moment to stabilize
                                            Thread.Sleep(10);
                                        }
                                        
                                        shapeList.Add(shape);
                                        
                                        try
                                        {
                                            Marshal.AddRef(Marshal.GetIUnknownForObject(shape));
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Warning: Could not add ref to shape: {ex.Message}");
                                        }
                                        
                                        ComReleaser.TrackObject(shape);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Error processing shape at index {i}: {ex.Message}");
                                }
                            }
                            
                            Console.WriteLine($"Successfully tracked {shapeCount} shapes");
                            GC.KeepAlive(shapeList);
                            GC.KeepAlive(shapes);
                            GC.KeepAlive(slide);
                            GC.KeepAlive(presentation);
                            GC.KeepAlive(pptApp);
                            
                            // Force a small GC to release any temporary objects
                            GC.Collect(0);
                            Thread.Sleep(100);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not access or process slide shapes: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not track slide shapes: {ex.Message}");
                    }
                    
                    // Save the presentation
                    string singleSlideOutputPath = outputPath;
                    if (!outputPath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                    {
                        singleSlideOutputPath = Path.ChangeExtension(outputPath, ".pptx");
                    }
                    
                    Console.WriteLine($"Preparing to save at: {singleSlideOutputPath}");
                    Console.WriteLine($"Directory exists: {Directory.Exists(Path.GetDirectoryName(singleSlideOutputPath))}");
                    
                    // Keep strong references during save
                    GC.KeepAlive(slide);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                    
                    try
                    {
                        Console.WriteLine("Starting save operation - maintaining all COM references...");
                        
                        // Brief pause before save to let COM objects stabilize
                        Thread.Sleep(500);
                        GC.KeepAlive(slide);
                        GC.KeepAlive(presentation);
                        GC.KeepAlive(pptApp);
                        
                        // Save the presentation
                        PowerPointOperations.SavePresentation(presentation, singleSlideOutputPath);
                        Console.WriteLine($"Presentation with slide #{slideIndex} saved to: {singleSlideOutputPath}");
                        
                        // Keep references alive after save
                        GC.KeepAlive(slide);
                        GC.KeepAlive(presentation);
                        GC.KeepAlive(pptApp);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error saving presentation: {ex.Message}");
                        if (ex.InnerException != null)
                        {
                            Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                        }
                    }
                    
                    // Delay cleanup for a moment to ensure save completes
                    Console.WriteLine("Delaying final cleanup to ensure save completes...");
                    Thread.Sleep(2000);
                    
                    // Keep references alive after delay
                    GC.KeepAlive(slide);
                    GC.KeepAlive(presentation);
                    GC.KeepAlive(pptApp);
                    
                    // Open the presentation
                    try
                    {
                        Console.WriteLine("Opening the presentation for review...");
                        System.Diagnostics.Process.Start(singleSlideOutputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not open the presentation: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to create slide #{slideIndex}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating single slide: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Cleanup in reverse order of creation
                try
                {
                    // Force all objects to stay alive until cleanup starts
                    if (slide != null) GC.KeepAlive(slide);
                    if (presentation != null) GC.KeepAlive(presentation);
                    if (pptApp != null) GC.KeepAlive(pptApp);
                    
                    // Now perform cleanup
                    Console.WriteLine("Performing final cleanup...");
                    
                    // Release all tracked COM objects
                    Console.WriteLine("Releasing all tracked COM objects (count: " + ComReleaser.GetTrackedObjectCount() + ")");
                    ComReleaser.ReleaseAllTrackedObjects();
                    
                    // Close and release presentation
                    if (presentation != null)
                    {
                        try
                        {
                            presentation.Close();
                            Thread.Sleep(500);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error closing presentation: {ex.Message}");
                        }
                        
                        try
                        {
                            Marshal.ReleaseComObject(presentation);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error releasing presentation COM object: {ex.Message}");
                        }
                        
                        presentation = null;
                    }
                    
                    // Quit and release application
                    if (pptApp != null)
                    {
                        try
                        {
                            pptApp.Quit();
                            Thread.Sleep(500);
                        }
                        catch (Exception ex) 
                        {
                            Console.WriteLine($"Error quitting PowerPoint: {ex.Message}");
                        }
                        
                        try
                        {
                            Marshal.ReleaseComObject(pptApp);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error releasing PowerPoint COM object: {ex.Message}");
                        }
                        
                        pptApp = null;
                    }
                    
                    // Force garbage collection once
                    GC.Collect(2, GCCollectionMode.Forced, true, true);
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    // Verify PowerPoint processes are terminated
                    PowerPointOperations.TerminatePowerPointProcesses();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during cleanup: {ex.Message}");
                }
            }
        }
    }
}