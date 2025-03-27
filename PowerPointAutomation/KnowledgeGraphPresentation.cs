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
                // Initialize PowerPoint application
                InitializePowerPoint();

                // Apply custom theme
                ApplyCustomTheme();

                // Get slide layouts
                SetupSlideLayouts();

                // Create slides
                CreateTitleSlide();
                CreateIntroductionSlide();
                CreateCoreComponentsSlide();
                CreateStructuralExampleSlide();
                CreateTheoreticalFoundationsSlide();
                CreateImplementationTechnologiesSlide();
                CreateConstructionApproachesSlide();
                CreateMachineLearningIntegrationSlide();
                CreateApplicationsUseCasesSlide();
                CreateAdvantagesChallengesSlide();
                CreateFutureDirectionsSlide();
                CreateConclusionSlide();

                // Add transitions between slides
                AddSlideTransitions();

                // Add footer to all slides
                AddFooterToAllSlides();

                // Save the presentation
                presentation.SaveAs(outputPath);

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

            // Set theme colors using proper method call syntax rather than indexer
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorText).RGB = ColorTranslator.ToOle(primaryColor);     // Text/Background dark
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorBackground).RGB = ColorTranslator.ToOle(Color.White);      // Text/Background light
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorAccent1).RGB = ColorTranslator.ToOle(secondaryColor);   // Accent 1
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorAccent2).RGB = ColorTranslator.ToOle(accentColor);      // Accent 2
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorAccent3).RGB = ColorTranslator.ToOle(Color.FromArgb(146, 208, 80));  // Accent 3 - Green
            master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorAccent4).RGB = ColorTranslator.ToOle(Color.FromArgb(0, 176, 240));   // Accent 4 - Light blue

            // Set default font for the presentation using the Name property directly
            master.Theme.ThemeFontScheme.MajorFont.Name = "Segoe UI";
            master.Theme.ThemeFontScheme.MinorFont.Name = "Segoe UI";
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
        /// Creates a slide about applications and use cases of knowledge graphs
        /// </summary>
        private void CreateApplicationsUseCasesSlide()
        {
            Console.WriteLine("Creating applications and use cases slide...");

            var contentSlideGenerator = new ContentSlide(presentation, contentLayout);

            var slide = contentSlideGenerator.Generate(
                "Applications & Use Cases",
                new string[] {
                    "Enterprise Knowledge Management",
                    "� Corporate memory, expertise location, document management",
                    "Search and Recommendation Systems",
                    "� Semantic search, context-aware recommendations, knowledge panels",
                    "Research and Discovery",
                    "� Scientific literature analysis, drug discovery, patent analysis",
                    "Customer Intelligence",
                    "� 360� customer view, journey mapping, nuanced segmentation",
                    "Compliance and Risk Management",
                    "� Regulatory compliance, fraud detection, anti-money laundering"
                },
                "This slide showcases diverse applications of knowledge graphs across domains. " +
                "The hierarchical structure organizes use cases by industry or function. The " +
                "progressive disclosure through animation helps the audience focus on one " +
                "application area at a time.",
                true // Enable animations
            );

            // Add an SmartArt diagram to illustrate the use cases using numeric value
            var chart = slide.Shapes.AddSmartArt(
                (SmartArtLayout)1, // Use cycle layout by numeric value to avoid enum compatibility issues
                slide.Design.SlideMaster.Width - 350, // Right side
                240, // Y position
                300, // Width
                300  // Height
            );

            // Get the SmartArt nodes and customize them
            if (chart.SmartArt != null)
            {
                chart.SmartArt.AllNodes[0].TextFrame2.TextRange.Text = "Knowledge Graphs";
                chart.SmartArt.AllNodes[1].TextFrame2.TextRange.Text = "Enterprise";
                chart.SmartArt.AllNodes[2].TextFrame2.TextRange.Text = "Search";
                chart.SmartArt.AllNodes[3].TextFrame2.TextRange.Text = "Research";
                chart.SmartArt.AllNodes[4].TextFrame2.TextRange.Text = "Customer";
                chart.SmartArt.AllNodes[5].TextFrame2.TextRange.Text = "Compliance";

                // Add animation to the SmartArt
                var effect = slide.TimeLine.MainSequence.AddEffect(
                    chart,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            }
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
        /// Properly releases COM objects to prevent memory leaks
        /// </summary>
        private void CleanupComObjects()
        {
            Console.WriteLine("Cleaning up COM objects...");

            // Use the utility to safely release COM objects
            if (presentation != null)
            {
                Marshal.ReleaseComObject(presentation);
                presentation = null;
            }

            if (pptApp != null)
            {
                pptApp.Quit();
                Marshal.ReleaseComObject(pptApp);
                pptApp = null;
            }

            // Force garbage collection to clean up any remaining COM objects
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}