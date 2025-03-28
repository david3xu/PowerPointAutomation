using System;
using System.IO;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAutomation.Slides;
using PowerPointAutomation.Models;
using PowerPointAutomation.Utilities;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace PowerPointAutomation
{
    /// <summary>
    /// Main class responsible for generating a comprehensive knowledge graph presentation
    /// </summary>
    public class KnowledgeGraphPresentation
    {
        // Constants for PpSlideLayout values that might be missing in the current environment
        private const PpSlideLayout ppLayoutTitleAndContent = (PpSlideLayout)8;
        private const PpSlideLayout ppLayoutBlank = (PpSlideLayout)11;
        private const PpSlideLayout ppLayoutTwoObjectsAndText = (PpSlideLayout)3;
        
        // Interop objects that need to be tracked for cleanup
        private Application pptApp;
        private Presentation presentation;

        // CustomLayouts for different slide types
        private CustomLayout titleLayout;
        private CustomLayout contentLayout;
        private CustomLayout twoColumnLayout;
        private CustomLayout diagramLayout;
        private CustomLayout conclusionLayout;

        // Theme colors (for consistent branding)
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange
        private readonly Color lightColor = Color.FromArgb(242, 242, 242);    // Light gray

        // Slide generators
        private TitleSlide titleSlideGenerator;
        private IntroductionSlide introSlideGenerator;
        private CoreFeatureSlide featureSlideGenerator;
        private DiagramSlide diagramSlideGenerator;
        private ListSlide listSlideGenerator;
        private ComparisonSlide comparisonSlideGenerator;
        private SummarySlide summarySlideGenerator;

        // Output path for the presentation
        private string outputPath;

        /// <summary>
        /// Gets or sets the path where the presentation will be saved
        /// </summary>
        public string OutputPath
        {
            get { return outputPath; }
            set { outputPath = value; }
        }

        /// <summary>
        /// Initializes a new instance of the KnowledgeGraphPresentation class
        /// </summary>
        public KnowledgeGraphPresentation()
        {
            // Get project directory path - go from bin\Debug to the solution root
            string projectDir = AppDomain.CurrentDomain.BaseDirectory;
            // Navigate up from bin/Debug to the project directory
            for (int i = 0; i < 2; i++)
            {
                projectDir = Path.GetDirectoryName(projectDir);
            }
            
            // Get solution directory (one level up from project)
            string solutionDir = Path.GetDirectoryName(projectDir);
            
            // Initialize with default path if none provided
            string outputDir = Path.Combine(solutionDir, "PowerPointAutomation", "docs", "output");
            
            // If the project directory structure doesn't exist, fall back to desktop
            if (!Directory.Exists(outputDir))
            {
                outputPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "KnowledgeGraphPresentation.pptx");
            }
            else
            {
                outputPath = Path.Combine(outputDir, "KnowledgeGraphPresentation.pptx");
            }
        }

        /// <summary>
        /// Generates a complete knowledge graph presentation
        /// </summary>
        public void Generate()
        {
            try
            {
                Console.WriteLine("Creating Knowledge Graph presentation...");
                
                // Open PowerPoint and create a new presentation
                pptApp = new Application();
                ComReleaser.TrackObject(pptApp);
                
                presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);
                ComReleaser.TrackObject(presentation);
                
                // Force the presentation to be visible after creation for better troubleshooting
                pptApp.Visible = MsoTriState.msoTrue;
                
                // Configure presentation properties
                SetupPresentationProperties();
                
                Console.WriteLine("Setting up slide layouts...");
                
                // Declare layouts at this scope so they're available throughout the method  
                CustomLayout titleLayout = null;
                CustomLayout titleAndContentLayout = null;
                CustomLayout blankLayout = null;
                CustomLayout twoContentLayout = null;
                
                try 
                {
                    // Get the layout count to know what's available
                    int layoutCount = presentation.SlideMaster.CustomLayouts.Count;
                    Console.WriteLine($"Available layouts: {layoutCount}");
                    
                    // Find title layout (1-indexed)
                    if (layoutCount >= 1)
                    {
                        titleLayout = presentation.SlideMaster.CustomLayouts[1];
                        ComReleaser.TrackObject(titleLayout);
                    }
                    else
                    {
                        // Extremely unlikely case, but handle it
                        Console.WriteLine("Warning: No layouts available! Creating a default presentation.");
                        return; // Exit the method, can't continue without layouts
                    }
                    
                    // Find title and content layout
                    if (layoutCount >= 2) 
                    {
                        titleAndContentLayout = presentation.SlideMaster.CustomLayouts[2];
                        ComReleaser.TrackObject(titleAndContentLayout);
                    }
                    else
                    {
                        Console.WriteLine("Warning: Could not find title and content layout, using title layout");
                        titleAndContentLayout = titleLayout;
                    }
                    
                    // Find blank layout - don't try to access index 11 directly
                    if (layoutCount >= 3)
                    {
                        // Use layout 3 as a safer alternative to blank
                        blankLayout = presentation.SlideMaster.CustomLayouts[3];
                        ComReleaser.TrackObject(blankLayout);
                    }
                    else
                    {
                        Console.WriteLine("Warning: Could not find blank layout, using title layout");
                        blankLayout = titleLayout;
                    }
                    
                    // Find two content layout
                    if (layoutCount >= 4)
                    {
                        twoContentLayout = presentation.SlideMaster.CustomLayouts[4];
                        ComReleaser.TrackObject(twoContentLayout);
                    }
                    else
                    {
                        Console.WriteLine("Warning: Could not find two content layout, using title and content layout");
                        twoContentLayout = titleAndContentLayout;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Error setting up layouts: {ex.Message}");
                    Console.WriteLine("Falling back to default layouts");
                    
                    try {
                        // Try to get at least one layout, preferably index 1
                        titleLayout = presentation.SlideMaster.CustomLayouts[1];
                        ComReleaser.TrackObject(titleLayout);
                        
                        // Use the same layout for all
                        titleAndContentLayout = titleLayout;
                        blankLayout = titleLayout;
                        twoContentLayout = titleLayout;
                    }
                    catch {
                        // If even that fails, we can't continue
                        Console.WriteLine("Fatal error: Cannot create any layouts. Exiting presentation generation.");
                        return;
                    }
                }
                
                // Release some initial objects
                ComReleaser.ReleaseOldestObjects(10);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create title slide
                Console.WriteLine("Creating title slide...");
                titleSlideGenerator = new TitleSlide(presentation, titleLayout);
                Slide slide1 = titleSlideGenerator.Generate(
                    "Knowledge Graphs",
                    "Understanding the Foundation of Semantic Data",
                    "Dr. Jane Smith",
                    "AI Research Institute");
                ComReleaser.TrackObject(slide1);
                
                // Aggressive cleanup after title slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create introduction slide
                Console.WriteLine("Creating introduction slide...");
                introSlideGenerator = new IntroductionSlide(presentation, titleAndContentLayout);
                Slide slide2 = introSlideGenerator.Generate(
                    "What is a Knowledge Graph?",
                    "A knowledge graph is a network of entities, their semantic types, properties, and relationships. " +
                    "It integrates data from multiple sources and enables machines to understand the semantics (meaning) " +
                    "of data in a way that's closer to human understanding.\n\n" +
                    "Knowledge graphs form the foundation of many modern AI systems, including search engines, " +
                    "virtual assistants, recommendation systems, and data integration platforms.");
                ComReleaser.TrackObject(slide2);
                
                // Aggressive cleanup after introduction slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create core components slide
                Console.WriteLine("Creating core components slide...");
                featureSlideGenerator = new CoreFeatureSlide(presentation, titleAndContentLayout);
                Slide slide3 = featureSlideGenerator.Generate(
                    "Core Components",
                    new string[] {
                        "Entities (Nodes): Real-world objects, concepts, or events represented in the graph",
                        "Relationships (Edges): Connections between entities that describe how they relate to each other",
                        "Properties: Attributes that describe entities or relationships",
                        "Ontologies: Formal definitions of types, properties, and relationships",
                        "Inference Rules: Logical rules that derive new facts from existing ones"
                    });
                ComReleaser.TrackObject(slide3);
                
                // Aggressive cleanup after core components slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create structural example slide
                Console.WriteLine("Creating structural example slide...");
                diagramSlideGenerator = new DiagramSlide(presentation, blankLayout);
                Slide slide4 = diagramSlideGenerator.GenerateKnowledgeGraphDiagram(
                    "Structural Example",
                    "A simple knowledge graph showing entities and relationships");
                ComReleaser.TrackObject(slide4);
                
                // Aggressive cleanup after diagram slide creation
                ComReleaser.ReleaseOldestObjects(100);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create applications slide
                Console.WriteLine("Creating applications slide...");
                listSlideGenerator = new ListSlide(presentation, titleAndContentLayout);
                Slide slide5 = listSlideGenerator.GenerateBulletedList(
                    "Applications of Knowledge Graphs",
                    new string[] {
                        "Semantic Search: Understanding the intent and contextual meaning behind queries",
                        "Question Answering: Providing direct answers from structured knowledge",
                        "Recommendation Systems: Suggesting products, content, or connections based on relationships",
                        "Data Integration: Combining heterogeneous data sources with a unified semantic model",
                        "Explainable AI: Adding interpretability to AI systems through knowledge representation"
                    });
                ComReleaser.TrackObject(slide5);
                
                // Aggressive cleanup after applications slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create future directions slide
                Console.WriteLine("Creating future directions slide...");
                comparisonSlideGenerator = new ComparisonSlide(presentation, twoContentLayout);
                Slide slide6 = comparisonSlideGenerator.Generate(
                    "Future Directions",
                    new string[] {
                        "Multimodal Knowledge Graphs",
                        "Temporal Knowledge Representation",
                        "Federated Knowledge Graphs",
                        "Automated Knowledge Extraction",
                        "Quantum Knowledge Representations"
                    },
                    new string[] {
                        "Integration with Large Language Models",
                        "Zero-shot Knowledge Transfer",
                        "Neuro-symbolic AI Integration",
                        "Edge Computing Applications",
                        "Privacy-preserving Knowledge Sharing"
                    });
                ComReleaser.TrackObject(slide6);
                
                // Aggressive cleanup after future directions slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Create conclusion slide
                Console.WriteLine("Creating conclusion slide...");
                summarySlideGenerator = new SummarySlide(presentation, titleAndContentLayout);
                Slide slide7 = summarySlideGenerator.Generate(
                    "Conclusion",
                    "Knowledge graphs represent a powerful approach for structuring and reasoning with data. They enable:\n\n" +
                    "• Semantic understanding of information\n" +
                    "• Integration of heterogeneous data sources\n" +
                    "• Inference of new knowledge\n" +
                    "• More human-like AI capabilities\n\n" +
                    "As AI continues to evolve, knowledge graphs will play an increasingly important role in creating systems that can reason with data in context.",
                    "Contact: research@aigraphs.org");
                ComReleaser.TrackObject(slide7);
                
                // Aggressive cleanup after conclusion slide creation
                ComReleaser.ReleaseOldestObjects(50);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Add slide transitions for all slides
                Console.WriteLine("Adding slide transitions...");
                for (int i = 1; i <= presentation.Slides.Count; i++)
                {
                    try
                    {
                        Slide slide = presentation.Slides[i];
                        slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFade;
                        slide.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedMedium;
                        
                        // Release after each transition to prevent memory buildup
                        if (i % 2 == 0)
                        {
                            ComReleaser.ReleaseOldestObjects(10);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not set transition for slide {i}: {ex.Message}");
                    }
                }
                
                // Cleanup after transitions
                ComReleaser.ReleaseOldestObjects(20);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Add footer with slide numbers to all slides except title
                Console.WriteLine("Adding footers...");
                for (int i = 2; i <= presentation.Slides.Count; i++)
                {
                    try
                    {
                        presentation.Slides[i].HeadersFooters.SlideNumber.Visible = MsoTriState.msoTrue;
                        
                        // Release after each footer to prevent memory buildup
                        if (i % 3 == 0)
                        {
                            ComReleaser.ReleaseOldestObjects(10);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not set footer for slide {i}: {ex.Message}");
                    }
                }
                
                // Cleanup after footers
                ComReleaser.ReleaseOldestObjects(20);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Save the presentation
                Console.WriteLine($"Saving presentation to {this.outputPath}...");
                
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(this.outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                presentation.SaveAs(this.outputPath);
                Console.WriteLine("Presentation saved successfully.");
                
                // Display the presentation
                try
                {
                    Console.WriteLine("Opening presentation for review...");
                    System.Diagnostics.Process.Start(this.outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not open presentation for review: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating knowledge graph presentation: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                // Ensure all COM objects are cleaned up
                CleanupComObjects();
            }
        }

        /// <summary>
        /// Initializes the PowerPoint application instance
        /// </summary>
        private void InitializePowerPoint()
        {
            Console.WriteLine("Initializing PowerPoint...");

            // Create PowerPoint application
            pptApp = new Application();

            // Make PowerPoint visible during development for debugging
            // Set to False for production to run in background
            pptApp.Visible = MsoTriState.msoTrue;

            // Create new presentation
            presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);

            // Set presentation properties
            presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x9;
        }

        /// <summary>
        /// Applies a custom theme to the presentation
        /// </summary>
        private void ApplyCustomTheme()
        {
            Console.WriteLine("Applying custom theme...");

            // Get the first slide master using Office 2016+ syntax
            Master master = presentation.Designs[1].SlideMaster;

            // Set background color
            master.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);

            // Set theme colors using the compatibility layer instead of direct method calls
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 1, ColorTranslator.ToOle(primaryColor));     // Text/Background dark (index 1)
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 2, ColorTranslator.ToOle(Color.White));      // Text/Background light (index 2)
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 5, ColorTranslator.ToOle(secondaryColor));   // Accent 1 (index 5)
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 6, ColorTranslator.ToOle(accentColor));      // Accent 2 (index 6)
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 7, ColorTranslator.ToOle(Color.FromArgb(146, 208, 80)));  // Accent 3 (index 7)
            OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 8, ColorTranslator.ToOle(Color.FromArgb(0, 176, 240)));   // Accent 4 (index 8)

            // Set default font for the presentation using the compatibility layer
            OfficeCompatibility.SetThemeFont(master.Theme.ThemeFontScheme.MajorFont, "Segoe UI");
            OfficeCompatibility.SetThemeFont(master.Theme.ThemeFontScheme.MinorFont, "Segoe UI");
        }

        /// <summary>
        /// Sets up the slide layouts for different types of content
        /// </summary>
        private void SetupSlideLayouts()
        {
            Console.WriteLine("Setting up slide layouts...");

            // Get the first slide master using Office 2016+ syntax
            Master master = presentation.Designs[1].SlideMaster;

            // Store default layouts for different slide types using indexer syntax
            titleLayout = master.CustomLayouts[1];      // Title slide layout
            contentLayout = master.CustomLayouts[2];    // Title and content layout
            twoColumnLayout = master.CustomLayouts[3];  // Two content layout
            diagramLayout = master.CustomLayouts[7];    // Title and diagram layout
            conclusionLayout = master.CustomLayouts[2]; // Conclusion layout (reusing content layout)
        }

        /// <summary>
        /// Sets up the presentation properties
        /// </summary>
        private void SetupPresentationProperties()
        {
            // Set presentation properties
            presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x9;
        }

        /// <summary>
        /// Cleans up COM objects to prevent memory leaks and orphaned processes
        /// </summary>
        private void CleanupComObjects()
        {
            Console.WriteLine("Starting final cleanup process...");
            
            try
            {
                // First release slide-specific objects that might be holding references
                titleSlideGenerator = null;
                introSlideGenerator = null;
                featureSlideGenerator = null;
                diagramSlideGenerator = null;
                listSlideGenerator = null;
                comparisonSlideGenerator = null;
                summarySlideGenerator = null;
                
                // Force immediate garbage collection to clean up these references
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Close presentation and release
                if (presentation != null)
                {
                    try
                    {
                        // Try to close the presentation
                        presentation.Close();
                        Console.WriteLine("Presentation closed successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error closing presentation: {ex.Message}");
                    }
                    
                    // Release presentation COM object and nullify reference
                    Console.WriteLine("Releasing presentation COM object...");
                    object presObj = presentation;
                    presentation = null;
                    ComReleaser.ReleaseCOMObject(ref presObj);
                    
                    // Force immediate garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                // Quit PowerPoint application
                if (pptApp != null)
                {
                    try
                    {
                        // Try to quit the application
                        Console.WriteLine("Quitting PowerPoint application...");
                        pptApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error quitting PowerPoint: {ex.Message}");
                    }
                    
                    // Release PowerPoint COM object and nullify reference
                    Console.WriteLine("Releasing PowerPoint application COM object...");
                    object appObj = pptApp;
                    pptApp = null;
                    ComReleaser.ReleaseCOMObject(ref appObj);
                    
                    // Force immediate garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during primary COM object cleanup: {ex.Message}");
            }
            
            try
            {
                // Release remaining tracked COM objects in smaller batches
                Console.WriteLine("Releasing remaining tracked COM objects in small batches...");
                
                // First batch - release 10 objects at a time
                Console.WriteLine("First batch cleanup...");
                ComReleaser.ReleaseAllTrackedObjects(10);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Second batch - release 5 objects at a time for more thorough cleanup
                Console.WriteLine("Second batch cleanup (smaller batch size)...");
                ComReleaser.ReleaseAllTrackedObjects(5);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Final sweep - release any remaining objects one by one
                Console.WriteLine("Final sweep cleanup (one by one)...");
                ComReleaser.ReleaseAllTrackedObjects(1);
                
                // Force final garbage collection with maximum generations
                Console.WriteLine("Performing final garbage collection...");
                ComReleaser.FinalCleanup();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during batch COM object cleanup: {ex.Message}");
            }
                
            // Check if PowerPoint process is still running and kill if necessary
            try
            {
                // First check
                if (ComReleaser.IsProcessRunning("POWERPNT"))
                {
                    Console.WriteLine("Warning: PowerPoint process is still running. Attempting to terminate...");
                    int killed = ComReleaser.KillProcess("POWERPNT");
                    Console.WriteLine($"Terminated {killed} PowerPoint process(es).");
                    
                    // Wait a moment and check again
                    System.Threading.Thread.Sleep(500);
                    
                    // Second check after a brief pause
                    if (ComReleaser.IsProcessRunning("POWERPNT"))
                    {
                        Console.WriteLine("PowerPoint process still detected. Final termination attempt...");
                        killed = ComReleaser.KillProcess("POWERPNT");
                        Console.WriteLine($"Terminated {killed} remaining PowerPoint process(es).");
                    }
                }
                else
                {
                    Console.WriteLine("No remaining PowerPoint processes detected.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking/killing PowerPoint process: {ex.Message}");
            }
            
            Console.WriteLine("Cleanup complete.");
        }
    }
}