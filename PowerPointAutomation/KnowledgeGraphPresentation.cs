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

        /// <summary>
        /// Generates a complete knowledge graph presentation and saves it to the specified path
        /// </summary>
        /// <param name="outputPath">The file path where the presentation will be saved</param>
        public void Generate(string outputPath)
        {
            try
            {
                // Log Office version for diagnostic purposes
                Console.WriteLine($"Detected Office version: {OfficeCompatibility.GetOfficeVersion()}");

                // Initialize PowerPoint application with operation logging
                OfficeCompatibility.LogOperation("Initialize PowerPoint", () => InitializePowerPoint());

                // Apply custom theme
                OfficeCompatibility.LogOperation("Apply Custom Theme", () => ApplyCustomTheme());

                // Get slide layouts
                OfficeCompatibility.LogOperation("Setup Slide Layouts", () => SetupSlideLayouts());

                // Create slides with operation logging
                OfficeCompatibility.LogOperation("Create Title Slide", () => CreateTitleSlide());
                OfficeCompatibility.LogOperation("Create Introduction Slide", () => CreateIntroductionSlide());
                OfficeCompatibility.LogOperation("Create Core Components Slide", () => CreateCoreComponentsSlide());
                OfficeCompatibility.LogOperation("Create Structural Example Slide", () => CreateStructuralExampleSlide());
                OfficeCompatibility.LogOperation("Create Theoretical Foundations Slide", () => CreateTheoreticalFoundationsSlide());
                OfficeCompatibility.LogOperation("Create Implementation Technologies Slide", () => CreateImplementationTechnologiesSlide());
                OfficeCompatibility.LogOperation("Create Construction Approaches Slide", () => CreateConstructionApproachesSlide());
                OfficeCompatibility.LogOperation("Create Machine Learning Integration Slide", () => CreateMachineLearningIntegrationSlide());
                OfficeCompatibility.LogOperation("Create Applications UseCases Slide", () => CreateApplicationsUseCasesSlide());
                OfficeCompatibility.LogOperation("Create Advantages Challenges Slide", () => CreateAdvantagesChallengesSlide());
                OfficeCompatibility.LogOperation("Create Future Directions Slide", () => CreateFutureDirectionsSlide());
                OfficeCompatibility.LogOperation("Create Conclusion Slide", () => CreateConclusionSlide());

                // Add transitions between slides
                OfficeCompatibility.LogOperation("Add Slide Transitions", () => AddSlideTransitions());

                // Add footer to all slides
                OfficeCompatibility.LogOperation("Add Footer To All Slides", () => AddFooterToAllSlides());

                // Save the presentation
                OfficeCompatibility.LogOperation("Save Presentation", () => presentation.SaveAs(outputPath));

                Console.WriteLine("Presentation created successfully!");
            }
            finally
            {
                // Always clean up COM objects
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

        #region Slide Creation Methods

        /// <summary>
        /// Creates the title slide
        /// </summary>
        private void CreateTitleSlide()
        {
            Console.WriteLine("Creating title slide...");

            var titleSlideGenerator = new TitleSlide(presentation, titleLayout);

            titleSlideGenerator.Generate(
                "Knowledge Graphs",
                "A Comprehensive Introduction",
                "Presented by PowerPoint Automation",
                "The title slide introduces the presentation topic with a visually appealing layout. " +
                "The main title uses a larger font with the primary color, while the subtitle uses a " +
                "slightly smaller font with the secondary color. This establishes the visual hierarchy " +
                "that will be consistent throughout the presentation."
            );
        }

        /// <summary>
        /// Creates the introduction slide explaining what knowledge graphs are
        /// </summary>
        private void CreateIntroductionSlide()
        {
            Console.WriteLine("Creating introduction slide...");

            var contentSlideGenerator = new ContentSlide(presentation, contentLayout);

            contentSlideGenerator.Generate(
                "Introduction to Knowledge Graphs",
                new string[] {
                    "Knowledge graphs represent information as interconnected entities and relationships",
                    "Semantic networks that represent real-world entities (objects, events, concepts)",
                    "Bridge structured and unstructured data for human and machine interpretation",
                    "Enable sophisticated reasoning, discovery, and analysis capabilities",
                    "Create a flexible yet robust foundation for knowledge management"
                },
                "This slide introduces the fundamental concept of knowledge graphs as networks of " +
                "entities and relationships. Emphasize how they differ from traditional data structures " +
                "by explicitly modeling connections. The bullet points build progressively to highlight " +
                "the key characteristics of knowledge graphs."
            );
        }

        /// <summary>
        /// Creates a slide explaining the core components of knowledge graphs
        /// </summary>
        private void CreateCoreComponentsSlide()
        {
            Console.WriteLine("Creating core components slide...");

            var contentSlideGenerator = new ContentSlide(presentation, contentLayout);

            contentSlideGenerator.Generate(
                "Core Components of Knowledge Graphs",
                new string[] {
                    "Nodes (Entities): Discrete objects, concepts, events, or states",
                    "� Unique identifiers, categorized by type, contain properties",
                    "Edges (Relationships): Connect nodes and define how entities relate",
                    "� Directed connections with semantic meaning, typed, may contain properties",
                    "Labels and Properties: Provide additional context and attributes",
                    "� Node labels denote entity types, edge labels specify relationship types"
                },
                "This slide outlines the three fundamental building blocks of knowledge graphs. " +
                "The nested bullet points provide more detail about each component. The slide uses " +
                "progressive disclosure through animation to avoid overwhelming the audience with " +
                "too much information at once.",
                true // Enable animations for bullet points
            );
        }

        /// <summary>
        /// Creates a slide with a visual example of a knowledge graph structure
        /// </summary>
        private void CreateStructuralExampleSlide()
        {
            Console.WriteLine("Creating structural example slide...");

            var diagramSlideGenerator = new DiagramSlide(presentation, diagramLayout);

            diagramSlideGenerator.GenerateKnowledgeGraphDiagram(
                "Structural Example",
                "A simple knowledge graph fragment representing company information",
                "This slide presents a visual example of a knowledge graph structure. " +
                "The diagram shows how entities are connected through relationships, " +
                "with each having specific properties. The animation sequence reveals " +
                "the components step by step to help the audience understand how the " +
                "graph is constructed.",
                true // Enable animations
            );
        }

        /// <summary>
        /// Creates a slide about the theoretical foundations of knowledge graphs
        /// </summary>
        private void CreateTheoreticalFoundationsSlide()
        {
            Console.WriteLine("Creating theoretical foundations slide...");

            var contentSlideGenerator = new ContentSlide(presentation, twoColumnLayout);

            contentSlideGenerator.GenerateTwoColumn(
                "Theoretical Foundations",
                new string[] {
                    "Graph Theory",
                    "� Connectivity, centrality, community structure",
                    "� Path analysis, network algorithms",
                    "Semantic Networks",
                    "� Conceptual associations, hierarchical organizations",
                    "� Meaning representation"
                },
                new string[] {
                    "Ontological Modeling",
                    "� Class hierarchies, property definitions",
                    "� Axioms and rules, domain modeling",
                    "Knowledge Representation",
                    "� First-order logic, description logics",
                    "� Frame systems, semantic triples"
                },
                "This slide presents the theoretical foundations that knowledge graphs build upon. " +
                "The two-column layout helps organize related but distinct concepts. Each foundation " +
                "includes sub-bullets that highlight key aspects. This structure helps the audience " +
                "understand the multidisciplinary nature of knowledge graph technology.",
                true // Enable animations
            );
        }

        /// <summary>
        /// Creates a slide about implementation technologies for knowledge graphs
        /// </summary>
        private void CreateImplementationTechnologiesSlide()
        {
            Console.WriteLine("Creating implementation technologies slide...");

            var contentSlideGenerator = new ContentSlide(presentation, contentLayout);

            // Create slides for implementation technologies
            var slide = contentSlideGenerator.Generate(
                "Implementation Technologies",
                new string[] {
                    "Data Models",
                    "� RDF (Resource Description Framework)",
                    "� Property Graphs",
                    "� Hypergraphs",
                    "� Knowledge Graph Embeddings",
                    "Storage Solutions",
                    "� Native Graph Databases (Neo4j, TigerGraph)",
                    "� RDF Triple Stores (AllegroGraph, Stardog)",
                    "� Multi-Model Databases (ArangoDB, OrientDB)"
                },
                "This slide covers the various technologies used to implement knowledge graphs. " +
                "It presents both data models and storage solutions. The hierarchical structure " +
                "helps organize related concepts, while the alternating colors help distinguish " +
                "between main categories and specific examples.",
                true // Enable animations
            );

            // Add code snippet for SPARQL query example
            PowerPointShape codeBox = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width - 350, // Right side
                slide.Design.SlideMaster.Height - 200, // Bottom area
                300, // Width
                150  // Height
            );

            // Format code box
            codeBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(240, 240, 240));
            codeBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkGray);
            codeBox.Line.Weight = 1.0f;

            // Add SPARQL query example
            codeBox.TextFrame.TextRange.Text =
                "Example SPARQL Query:\n\n" +
                "SELECT ?person ?company\n" +
                "WHERE {\n" +
                "  ?company a :Company .\n" +
                "  ?person a :Person .\n" +
                "  ?company :employs ?person .\n" +
                "  ?person :hasExpertise \"KG\" .\n" +
                "}";

            codeBox.TextFrame.TextRange.Font.Name = "Consolas";
            codeBox.TextFrame.TextRange.Font.Size = 10;

            // Animate the code box
            var effect = slide.TimeLine.MainSequence.AddEffect(
                codeBox,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
        }

        /// <summary>
        /// Creates a slide about construction approaches for knowledge graphs
        /// </summary>
        private void CreateConstructionApproachesSlide()
        {
            Console.WriteLine("Creating construction approaches slide...");

            var contentSlideGenerator = new ContentSlide(presentation, twoColumnLayout);

            contentSlideGenerator.GenerateTwoColumn(
                "Construction Approaches",
                new string[] {
                    "Manual Curation",
                    "� Expert-driven construction ensuring high quality",
                    "� Time-intensive, difficult to scale",
                    "� Critical domains requiring accuracy (healthcare, legal)",
                    "Automated Extraction",
                    "� Information extraction from text",
                    "� Wrapper induction from web pages",
                    "� Database transformation from relational data"
                },
                new string[] {
                    "Hybrid Approaches",
                    "� Bootstrap and refine: Automated with manual verification",
                    "� Pattern-based expansion: Using patterns to extend examples",
                    "� Distant supervision: Leveraging existing knowledge",
                    "� Continuous feedback: Incorporating user corrections",
                    "Evaluation Criteria",
                    "� Accuracy, coverage, consistency",
                    "� Semantic validity, alignment with domain knowledge"
                },
                "This slide presents different methodologies for constructing knowledge graphs. " +
                "The two-column layout creates a natural comparison between approaches. The " +
                "progressive disclosure through animation helps maintain focus on one approach " +
                "at a time before revealing the next one.",
                true // Enable animations
            );
        }

        /// <summary>
        /// Creates a slide about machine learning integration with knowledge graphs
        /// </summary>
        private void CreateMachineLearningIntegrationSlide()
        {
            Console.WriteLine("Creating machine learning integration slide...");

            var diagramSlideGenerator = new DiagramSlide(presentation, diagramLayout);

            diagramSlideGenerator.GenerateMLIntegrationDiagram(
                "Machine Learning Integration",
                "How knowledge graphs and machine learning interact",
                "This slide visualizes the bidirectional relationship between knowledge graphs " +
                "and machine learning. The circular diagram shows how machine learning can help " +
                "build and enhance knowledge graphs, while knowledge graphs can improve machine " +
                "learning models through structured knowledge. The animation sequence reveals " +
                "these relationships step by step."
            );
        }

        /// <summary>
        /// Creates a slide about real-world applications and use cases
        /// </summary>
        private void CreateApplicationsUseCasesSlide()
        {
            Console.WriteLine("Creating applications and use cases slide...");

            // Add slide with title layout
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, contentLayout);

            // Set slide title
            slide.Shapes.Title.TextFrame.TextRange.Text = "Applications & Use Cases";
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add title underline
            PowerPointShape titleUnderline = slide.Shapes.AddLine(
                slide.Shapes.Title.Left,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5,
                slide.Shapes.Title.Left + slide.Shapes.Title.Width * 0.4f,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5
            );
            titleUnderline.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
            titleUnderline.Line.Weight = 3.0f;

            // Add introduction text
            PowerPointShape introTextBox = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                50, 100, 640, 80
            );
            introTextBox.TextFrame.TextRange.Text = "Knowledge graphs power a wide range of applications across industries:";
            introTextBox.TextFrame.TextRange.Font.Size = 20;
            introTextBox.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);

            // Calculate SmartArt position
            float smartArtLeft = 50;
            float smartArtTop = 180;
            float smartArtWidth = 640;
            float smartArtHeight = 300;

            // Get SmartArt layout using compatibility layer
            object smartArtLayout = OfficeCompatibility.GetSmartArtLayout(slide.Application, 1); // Cycle layout

            // Add SmartArt if layout is available
            if (smartArtLayout != null)
            {
                // Add SmartArt diagram using compatibility-safe approach
                var chart = slide.Shapes.AddSmartArt(
                    smartArtLayout, 
                    smartArtLeft, smartArtTop, smartArtWidth, smartArtHeight);

                // Get the SmartArt nodes and customize them
                if (chart.SmartArt != null)
                {
                    try
                    {
                        chart.SmartArt.AllNodes[0].TextFrame2.TextRange.Text = "Knowledge Graphs";
                        chart.SmartArt.AllNodes[1].TextFrame2.TextRange.Text = "Enterprise";
                        chart.SmartArt.AllNodes[2].TextFrame2.TextRange.Text = "Search";
                        chart.SmartArt.AllNodes[3].TextFrame2.TextRange.Text = "Research";
                        chart.SmartArt.AllNodes[4].TextFrame2.TextRange.Text = "Customer";
                        chart.SmartArt.AllNodes[5].TextFrame2.TextRange.Text = "Compliance";

                        // Add animation to the SmartArt
                        slide.TimeLine.MainSequence.AddEffect(
                            chart,
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelAllAtOnce,
                            MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    }
                    catch (Exception ex)
                    {
                        // Log SmartArt text setting error for debugging
                        Console.WriteLine($"Error setting SmartArt text: {ex.Message}");
                    }
                }
            }
            else
            {
                // Fallback: Create a simple shape layout instead of SmartArt
                PowerPointShape fallbackShape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    smartArtLeft + smartArtWidth/3, smartArtTop, smartArtWidth/3, 60);
                
                fallbackShape.TextFrame.TextRange.Text = "Knowledge Graphs";
                fallbackShape.TextFrame.TextRange.Font.Size = 24;
                fallbackShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                fallbackShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                fallbackShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(primaryColor);
                
                // Create circle shapes for use cases around the center box
                CreateUseCase(slide, "Enterprise", smartArtLeft + 60, smartArtTop + 100);
                CreateUseCase(slide, "Search", smartArtLeft + smartArtWidth - 150, smartArtTop + 100);
                CreateUseCase(slide, "Research", smartArtLeft + 60, smartArtTop + 200);
                CreateUseCase(slide, "Customer", smartArtLeft + smartArtWidth - 150, smartArtTop + 200);
                CreateUseCase(slide, "Compliance", smartArtLeft + smartArtWidth/2 - 45, smartArtTop + 250);
            }

            // Add speaker notes
            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text =
                "This slide highlights how knowledge graphs are applied across different domains. " +
                "The diagram shows various industries and applications. In enterprise settings, " +
                "knowledge graphs connect disparate data sources. For search, they enhance relevance " +
                "and context. Research applications leverage connections to find new insights. " +
                "Customer applications include recommendation engines and personalization. " +
                "Compliance applications use knowledge graphs for risk assessment and audit trails.";
        }
        
        /// <summary>
        /// Helper method to create a use case bubble (for SmartArt fallback)
        /// </summary>
        private void CreateUseCase(Slide slide, string text, float left, float top)
        {
            PowerPointShape circle = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                left, top, 90, 90);
                
            circle.TextFrame.TextRange.Text = text;
            circle.TextFrame.TextRange.Font.Size = 18;
            circle.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            circle.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            circle.Fill.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
            
            // Add connector line to main shape
            PowerPointShape connector = slide.Shapes.AddLine(
                left + 45, top, 
                left + 45, top - 20);
            connector.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
            connector.Line.Weight = 2.0f;
        }

        /// <summary>
        /// Creates a slide about advantages and challenges of knowledge graphs
        /// </summary>
        private void CreateAdvantagesChallengesSlide()
        {
            Console.WriteLine("Creating advantages and challenges slide...");

            var contentSlideGenerator = new ContentSlide(presentation, twoColumnLayout);

            contentSlideGenerator.GenerateTwoColumn(
                "Advantages & Challenges",
                new string[] {
                    "Key Advantages",
                    "� Contextual Understanding: Data with semantic context",
                    "� Flexibility: Adaptable to evolving information needs",
                    "� Integration Capability: Unifies diverse data sources",
                    "� Inferential Power: Discovers implicit knowledge",
                    "� Human-Interpretable: Aligns with conceptual understanding"
                },
                new string[] {
                    "Implementation Challenges",
                    "� Construction Complexity: Significant effort required",
                    "� Schema Evolution: Maintaining consistency while growing",
                    "� Performance at Scale: Optimizing for large graphs",
                    "� Quality Assurance: Ensuring accuracy across assertions",
                    "� User Adoption: Requiring new query paradigms"
                },
                "This slide presents a balanced view of both the advantages and challenges of " +
                "knowledge graph implementations. The side-by-side comparison helps decision-makers " +
                "understand both the benefits and potential obstacles. The contrasting colors " +
                "visually distinguish between advantages and challenges.",
                true // Enable animations
            );
        }

        /// <summary>
        /// Creates a slide about future directions in knowledge graph technology
        /// </summary>
        private void CreateFutureDirectionsSlide()
        {
            Console.WriteLine("Creating future directions slide...");

            var contentSlideGenerator = new ContentSlide(presentation, contentLayout);

            contentSlideGenerator.Generate(
                "Future Directions",
                new string[] {
                    "Self-Improving Knowledge Graphs",
                    "� Automated knowledge acquisition and contradiction detection",
                    "� Confidence scoring and active learning",
                    "Multimodal Knowledge Graphs",
                    "� Visual, temporal, spatial, and numerical integration",
                    "� Cross-modal reasoning and representation",
                    "Neuro-Symbolic Integration",
                    "� Combining neural networks with symbolic logic",
                    "� Using knowledge graphs to explain AI decisions",
                    "� Foundation model integration with knowledge graphs"
                },
                "This slide explores emerging trends and future developments in knowledge graph " +
                "technology. The hierarchical structure helps organize related concepts, while " +
                "the animation sequence creates a sense of progression from current capabilities " +
                "toward future innovations.",
                true // Enable animations
            );
        }

        /// <summary>
        /// Creates the conclusion slide summarizing key points
        /// </summary>
        private void CreateConclusionSlide()
        {
            Console.WriteLine("Creating conclusion slide...");

            var conclusionSlideGenerator = new ConclusionSlide(presentation, conclusionLayout);

            conclusionSlideGenerator.Generate(
                "Conclusion",
                "Knowledge graphs represent a transformative approach to information management, enabling organizations to move beyond data silos toward connected intelligence. By explicitly modeling relationships between entities, knowledge graphs provide context that traditional databases lack, supporting sophisticated reasoning and discovery.\n\n" +
                "While implementing knowledge graphs presents challenges in construction, maintenance, and scalability, the benefits of contextual understanding, flexible integration, and inferential capabilities make them increasingly essential for organizations dealing with complex, interconnected information.",
                "Thank you!",
                "contact@example.com",
                "This conclusion slide summarizes the key takeaways about knowledge graphs. " +
                "It reinforces the main value proposition while acknowledging the implementation " +
                "challenges. The call to action encourages the audience to consider how knowledge " +
                "graphs might apply to their specific context."
            );
        }

        #endregion

        /// <summary>
        /// Adds transitions between slides for a more polished presentation
        /// </summary>
        private void AddSlideTransitions()
        {
            Console.WriteLine("Adding slide transitions...");

            // Apply a simple fade transition to all slides
            for (int i = 1; i <= presentation.Slides.Count; i++)
            {
                Slide slide = presentation.Slides[i];

                // Apply a fade transition
                slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFade; // Use the enum value instead of int

                // Set transition speed
                slide.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedMedium;

                // Advance on click (not automatically)
                slide.SlideShowTransition.AdvanceOnTime = MsoTriState.msoFalse;
                slide.SlideShowTransition.AdvanceOnClick = MsoTriState.msoTrue;
            }
        }

        /// <summary>
        /// Adds a consistent footer to all slides
        /// </summary>
        private void AddFooterToAllSlides()
        {
            Console.WriteLine("Adding footers to slides...");

            // Set footer properties for all slides
            for (int i = 1; i <= presentation.Slides.Count; i++)
            {
                // Skip title slide (slide 1)
                if (i == 1) continue;

                Slide slide = presentation.Slides[i];

                // Enable footer
                slide.HeadersFooters.Footer.Visible = MsoTriState.msoTrue;
                slide.HeadersFooters.Footer.Text = "Knowledge Graph Presentation | " + DateTime.Now.ToString("MMMM yyyy");

                // Enable slide numbers
                slide.HeadersFooters.SlideNumber.Visible = MsoTriState.msoTrue;

                // Date is not needed as it's included in the footer text
                slide.HeadersFooters.DateAndTime.Visible = MsoTriState.msoFalse;
            }
        }

        /// <summary>
        /// Cleans up COM objects to prevent memory leaks and orphaned processes
        /// </summary>
        private void CleanupComObjects()
        {
            Console.WriteLine("Cleaning up COM objects...");

            // Perform cleanup in reverse order (most recently created objects first)
            try
            {
                // Close presentation without saving changes
                if (presentation != null)
                {
                    try
                    {
                        // Try to close the presentation without saving
                        presentation.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error closing presentation: {ex.Message}");
                    }
                    
                    // Release COM object
                    object presObj = presentation;
                    presentation = null;
                    ComReleaser.ReleaseCOMObject(ref presObj);
                }

                // Quit PowerPoint application
                if (pptApp != null)
                {
                    try
                    {
                        // Try to quit the application
                        pptApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error quitting PowerPoint: {ex.Message}");
                    }
                    
                    // Release COM object
                    object appObj = pptApp;
                    pptApp = null;
                    ComReleaser.ReleaseCOMObject(ref appObj);
                }
            }
            finally
            {
                // Release all other tracked COM objects
                ComReleaser.ReleaseAllTrackedObjects();
                
                // Force garbage collection
                ComReleaser.FinalCleanup();
                
                // Check if PowerPoint process is still running
                if (ComReleaser.IsProcessRunning("POWERPNT"))
                {
                    Console.WriteLine("Warning: PowerPoint process is still running. Attempting to terminate...");
                    int killed = ComReleaser.KillProcess("POWERPNT");
                    Console.WriteLine($"Terminated {killed} PowerPoint process(es).");
                }
            }
        }
    }
}