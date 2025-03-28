using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using PowerPointAutomation.Utilities;
using System.Collections.Generic;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class responsible for generating diagram slides with visualizations
    /// </summary>
    public class DiagramSlide
    {
        private Presentation presentation;
        private CustomLayout layout;

        // Theme colors for consistent branding
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange
        private readonly Color nodeColor = Color.FromArgb(91, 155, 213);      // Light blue for nodes
        private readonly Color edgeColor = Color.FromArgb(237, 125, 49);      // Orange for edges

        /// <summary>
        /// Initializes a new instance of the DiagramSlide class
        /// </summary>
        /// <param name="presentation">The PowerPoint presentation</param>
        /// <param name="layout">The slide layout to use</param>
        public DiagramSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a knowledge graph diagram slide
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">Optional subtitle</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <param name="animate">Whether to add animations</param>
        /// <returns>The created slide</returns>
        public Slide GenerateKnowledgeGraphDiagram(string title, string subtitle = null, string notes = null, bool animate = true)
        {
            Slide slide = null;
            try
            {
                // Add diagram slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);
                ComReleaser.TrackObject(slide);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(slide));

                // Create a custom title instead of trying to access the Title shape placeholder
                PowerPointShape titleShape = null;
                try
                {
                    titleShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        50, // Left
                        20, // Top
                        slide.Design.SlideMaster.Width - 100, // Width
                        50 // Height
                    );
                    ComReleaser.TrackObject(titleShape);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(titleShape));
                    
                    // Format the title
                    titleShape.TextFrame.TextRange.Text = title;
                    titleShape.TextFrame.TextRange.Font.Size = 36;
                    titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
                    titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                    titleShape.Line.Visible = MsoTriState.msoFalse;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not create title: {ex.Message}. Continuing...");
                }

                // Add subtitle if provided
                PowerPointShape subtitleShape = null;
                if (!string.IsNullOrEmpty(subtitle))
                {
                    try
                    {
                        float subtitleTop = titleShape != null ? titleShape.Top + titleShape.Height + 10 : 80;
                        subtitleShape = slide.Shapes.AddTextbox(
                            MsoTextOrientation.msoTextOrientationHorizontal,
                            titleShape != null ? titleShape.Left : 50,
                            subtitleTop,
                            slide.Design.SlideMaster.Width - (titleShape != null ? titleShape.Left * 2 : 100),
                            30
                        );
                        ComReleaser.TrackObject(subtitleShape);
                        
                        // Add manual ref to prevent RCW separation
                        Marshal.AddRef(Marshal.GetIUnknownForObject(subtitleShape));
                        
                        subtitleShape.TextFrame.TextRange.Text = subtitle;
                        subtitleShape.TextFrame.TextRange.Font.Size = 18;
                        subtitleShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                        subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not create subtitle: {ex.Message}. Continuing...");
                    }
                }

                // Track shape count to avoid exceeding PowerPoint's limits
                int maxShapesPerSlide = 25; // Conservative limit to prevent issues
                int currentShapeCount = slide.Shapes.Count;
                Console.WriteLine($"Current shape count: {currentShapeCount}");

                // Add a legend in the bottom right
                if (currentShapeCount + 2 < maxShapesPerSlide)
                {
                    try
                    {
                        AddLegendItem(slide, slide.Design.SlideMaster.Width - 200, slide.Design.SlideMaster.Height - 100, nodeColor, "Entity");
                        AddLegendItem(slide, slide.Design.SlideMaster.Width - 200, slide.Design.SlideMaster.Height - 70, edgeColor, "Relationship");
                        currentShapeCount += 2;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not add legend: {ex.Message}. Continuing without legend.");
                    }
                }

                // Create diagram elements - keep track of what we create for animations
                List<PowerPointShape> nodeShapes = new List<PowerPointShape>();
                List<PowerPointShape> edgeShapes = new List<PowerPointShape>();

                try
                {
                    // Create entity nodes
                    float centerX = slide.Design.SlideMaster.Width / 2;
                    float centerY = (slide.Design.SlideMaster.Height + 100) / 2;
                    
                    // Person node
                    if (currentShapeCount < maxShapesPerSlide)
                    {
                        try
                        {
                            PowerPointShape personNode = CreateEntityNode(slide, "Person", centerX - 200, centerY - 100, nodeColor);
                            if (personNode != null)
                            {
                                ComReleaser.TrackObject(personNode);
                                Marshal.AddRef(Marshal.GetIUnknownForObject(personNode));
                                nodeShapes.Add(personNode);
                                currentShapeCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not create Person node: {ex.Message}");
                        }
                    }
                    
                    // Organization node
                    if (currentShapeCount < maxShapesPerSlide)
                    {
                        try
                        {
                            PowerPointShape orgNode = CreateEntityNode(slide, "Organization", centerX + 200, centerY - 50, nodeColor);
                            if (orgNode != null)
                            {
                                ComReleaser.TrackObject(orgNode);
                                Marshal.AddRef(Marshal.GetIUnknownForObject(orgNode));
                                nodeShapes.Add(orgNode);
                                currentShapeCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not create Organization node: {ex.Message}");
                        }
                    }
                    
                    // Location node
                    if (currentShapeCount < maxShapesPerSlide)
                    {
                        try
                        {
                            PowerPointShape locationNode = CreateEntityNode(slide, "Location", centerX, centerY + 150, nodeColor);
                            if (locationNode != null)
                            {
                                ComReleaser.TrackObject(locationNode);
                                Marshal.AddRef(Marshal.GetIUnknownForObject(locationNode));
                                nodeShapes.Add(locationNode);
                                currentShapeCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not create Location node: {ex.Message}");
                        }
                    }
                    
                    // Create relationships safely
                    if (nodeShapes.Count >= 2)
                    {
                        // No more than 3 nodes can be created above, so these indices are safe
                        PowerPointShape personNode = nodeShapes.Count > 0 ? nodeShapes[0] : null;
                        PowerPointShape orgNode = nodeShapes.Count > 1 ? nodeShapes[1] : null;
                        PowerPointShape locationNode = nodeShapes.Count > 2 ? nodeShapes[2] : null;
                        
                        // Create relationships between entities
                        if (personNode != null && orgNode != null && currentShapeCount < maxShapesPerSlide)
                        {
                            try
                            {
                                PowerPointShape worksForEdge = CreateRelationship(slide, personNode, orgNode, "WORKS_FOR");
                                if (worksForEdge != null)
                                {
                                    ComReleaser.TrackObject(worksForEdge);
                                    Marshal.AddRef(Marshal.GetIUnknownForObject(worksForEdge));
                                    edgeShapes.Add(worksForEdge);
                                    currentShapeCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Could not create WORKS_FOR relationship: {ex.Message}");
                            }
                        }
                        
                        if (personNode != null && locationNode != null && currentShapeCount < maxShapesPerSlide)
                        {
                            try
                            {
                                PowerPointShape livesInEdge = CreateRelationship(slide, personNode, locationNode, "LIVES_IN");
                                if (livesInEdge != null)
                                {
                                    ComReleaser.TrackObject(livesInEdge);
                                    Marshal.AddRef(Marshal.GetIUnknownForObject(livesInEdge));
                                    edgeShapes.Add(livesInEdge);
                                    currentShapeCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Could not create LIVES_IN relationship: {ex.Message}");
                            }
                        }
                        
                        if (orgNode != null && locationNode != null && currentShapeCount < maxShapesPerSlide)
                        {
                            try
                            {
                                PowerPointShape locatedInEdge = CreateRelationship(slide, orgNode, locationNode, "LOCATED_IN");
                                if (locatedInEdge != null)
                                {
                                    ComReleaser.TrackObject(locatedInEdge);
                                    Marshal.AddRef(Marshal.GetIUnknownForObject(locatedInEdge));
                                    edgeShapes.Add(locatedInEdge);
                                    currentShapeCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Could not create LOCATED_IN relationship: {ex.Message}");
                            }
                        }
                    }
                    
                    Console.WriteLine($"Created diagram with {nodeShapes.Count} nodes and {edgeShapes.Count} edges");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Error creating graph elements: {ex.Message}. Continuing with partial diagram.");
                }

                // Add animations if requested
                if (animate && slide.TimeLine != null && nodeShapes.Count > 0)
                {
                    try
                    {
                        // Check timeLine and sequence access before proceeding
                        if (slide.TimeLine == null || slide.TimeLine.MainSequence == null)
                        {
                            Console.WriteLine("Cannot add animations: Timeline or MainSequence is null");
                        }
                        else
                        {
                            try 
                            {
                                int sequenceCount = slide.TimeLine.MainSequence.Count;
                                Console.WriteLine($"Current animation sequence count: {sequenceCount}");
                                
                                // Limit animation to prevent index out of range errors
                                int maxAnimations = 10; // Even more conservative than before
                                int nodesToAnimate = Math.Min(nodeShapes.Count, maxAnimations);
                                
                                Console.WriteLine($"Animating {nodesToAnimate} nodes out of {nodeShapes.Count} total");
                                
                                // Animate nodes first - with limit
                                for (int i = 0; i < nodesToAnimate; i++)
                                {
                                    try
                                    {
                                        Effect nodeEffect = slide.TimeLine.MainSequence.AddEffect(
                                            nodeShapes[i],
                                            GetNodeAnimationEffect(i),
                                            MsoAnimateByLevel.msoAnimateLevelNone,
                                            i == 0 ? MsoAnimTriggerType.msoAnimTriggerOnPageClick : MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                        ComReleaser.TrackObject(nodeEffect);

                                        nodeEffect.Timing.Duration = 0.5f;
                                        
                                        // Perform cleanup after every node to prevent memory buildup
                                        ComReleaser.ReleaseOldestObjects(3);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Warning: Could not add animation for node {i}: {ex.Message}");
                                    }
                                }
                                
                                // Check if we can add more animations
                                int remainingSlots = maxAnimations - nodesToAnimate;
                                int edgesToAnimate = Math.Min(edgeShapes.Count, remainingSlots);
                                
                                Console.WriteLine($"Animating {edgesToAnimate} edges out of {edgeShapes.Count} total");
                                
                                // Then animate edges
                                for (int i = 0; i < edgesToAnimate; i++)
                                {
                                    try
                                    {
                                        Effect edgeEffect = slide.TimeLine.MainSequence.AddEffect(
                                            edgeShapes[i],
                                            MsoAnimEffect.msoAnimEffectFade,
                                            MsoAnimateByLevel.msoAnimateLevelNone,
                                            MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                        ComReleaser.TrackObject(edgeEffect);

                                        edgeEffect.Timing.Duration = 0.4f;
                                        
                                        // Perform cleanup after every animation to prevent memory buildup
                                        ComReleaser.ReleaseOldestObjects(3);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Warning: Could not add animation for edge {i}: {ex.Message}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Error configuring animations: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not add animations to diagram: {ex.Message}. Continuing without animations.");
                    }
                }

                // Add speaker notes if provided
                if (!string.IsNullOrEmpty(notes))
                {
                    try
                    {
                        // Safely access notes page
                        if (slide.NotesPage != null && slide.NotesPage.Shapes.Count >= 2)
                        {
                            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
                        }
                        else
                        {
                            Console.WriteLine("Warning: Cannot add notes - notes shape not available");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not add notes to the slide: {ex.Message}");
                    }
                }

                // Final cleanup before returning
                try
                {
                    ComReleaser.ReleaseOldestObjects(10);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not release oldest objects: {ex.Message}");
                }
                
                return slide;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating knowledge graph diagram: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return slide; // Return the partially created slide or null
            }
        }

        /// <summary>
        /// Creates a diagram illustrating machine learning integration with knowledge graphs
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">Optional subtitle</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <returns>The created slide</returns>
        public Slide GenerateMLIntegrationDiagram(string title, string subtitle = null, string notes = null)
        {
            // Add diagram slide
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Create a custom title instead of trying to access the Title shape placeholder
            PowerPointShape titleShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                50, // Left
                20, // Top
                slide.Design.SlideMaster.Width - 100, // Width
                50 // Height
            );
            
            // Format the title
            titleShape.TextFrame.TextRange.Text = title;
            titleShape.TextFrame.TextRange.Font.Size = 36;
            titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
            titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
            titleShape.Line.Visible = MsoTriState.msoFalse;

            // Add subtitle if provided
            if (!string.IsNullOrEmpty(subtitle))
            {
                float subtitleTop = titleShape.Top + titleShape.Height + 10;
                PowerPointShape subtitleShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    titleShape.Left,
                    subtitleTop,
                    slide.Design.SlideMaster.Width - (titleShape.Left * 2),
                    30
                );

                subtitleShape.TextFrame.TextRange.Text = subtitle;
                subtitleShape.TextFrame.TextRange.Font.Size = 18;
                subtitleShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
            }

            // Create a circular diagram layout
            float centerX = slide.Design.SlideMaster.Width / 2;
            float centerY = (slide.Design.SlideMaster.Height + 100) / 2;
            float radius = 180;

            // Create the main center circle for Knowledge Graphs
            PowerPointShape kgCircle = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                centerX - 80,
                centerY - 80,
                160,
                160
            );
            ComReleaser.TrackObject(kgCircle);
            Marshal.AddRef(Marshal.GetIUnknownForObject(kgCircle));

            kgCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196)); // Blue
            kgCircle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125));
            kgCircle.Line.Weight = 2.0f;

            PowerPointShape kgText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - 70,
                centerY - 25,
                140,
                50
            );
            ComReleaser.TrackObject(kgText);
            Marshal.AddRef(Marshal.GetIUnknownForObject(kgText));

            kgText.TextFrame.TextRange.Text = "Knowledge Graphs";
            kgText.TextFrame.TextRange.Font.Size = 16;
            kgText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            kgText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            kgText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            kgText.Line.Visible = MsoTriState.msoFalse;

            // Create ML circle
            PowerPointShape mlCircle = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                centerX - radius - 60,
                centerY - 60,
                120,
                120
            );
            ComReleaser.TrackObject(mlCircle);
            Marshal.AddRef(Marshal.GetIUnknownForObject(mlCircle));

            mlCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(112, 173, 71)); // Green
            mlCircle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(84, 130, 53));
            mlCircle.Line.Weight = 2.0f;

            PowerPointShape mlText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius - 50,
                centerY - 20,
                100,
                50
            );
            ComReleaser.TrackObject(mlText);
            Marshal.AddRef(Marshal.GetIUnknownForObject(mlText));

            mlText.TextFrame.TextRange.Text = "Machine Learning";
            mlText.TextFrame.TextRange.Font.Size = 14;
            mlText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            mlText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            mlText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            mlText.Line.Visible = MsoTriState.msoFalse;

            // Create four integration point circles around the main KG circle
            string[] integrationPoints = {
                "Entity Recognition",
                "Relationship Extraction",
                "Knowledge Graph Completion",
                "Graph Neural Networks"
            };

            PowerPointShape[] integrationCircles = new PowerPointShape[integrationPoints.Length];
            List<PowerPointShape> integrationTexts = new List<PowerPointShape>();
            List<PowerPointShape> connectors = new List<PowerPointShape>();

            for (int i = 0; i < integrationPoints.Length; i++)
            {
                float angle = (float)(2 * Math.PI * i / integrationPoints.Length);
                float x = centerX + (float)(radius * Math.Cos(angle));
                float y = centerY + (float)(radius * Math.Sin(angle));

                integrationCircles[i] = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeOval,
                    x - 50,
                    y - 50,
                    100,
                    100
                );
                ComReleaser.TrackObject(integrationCircles[i]);
                Marshal.AddRef(Marshal.GetIUnknownForObject(integrationCircles[i]));

                integrationCircles[i].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(237, 125, 49)); // Orange
                integrationCircles[i].Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(191, 96, 27));
                integrationCircles[i].Line.Weight = 1.5f;

                PowerPointShape integrationText = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    x - 45,
                    y - 25,
                    90,
                    50
                );
                ComReleaser.TrackObject(integrationText);
                Marshal.AddRef(Marshal.GetIUnknownForObject(integrationText));
                integrationTexts.Add(integrationText);

                integrationText.TextFrame.TextRange.Text = integrationPoints[i];
                integrationText.TextFrame.TextRange.Font.Size = 11;
                integrationText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                integrationText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                integrationText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                integrationText.Line.Visible = MsoTriState.msoFalse;

                // Create connector line from KG to integration point
                PowerPointShape connector = slide.Shapes.AddConnector(
                    MsoConnectorType.msoConnectorStraight,
                    centerX,
                    centerY,
                    x,
                    y
                );
                ComReleaser.TrackObject(connector);
                Marshal.AddRef(Marshal.GetIUnknownForObject(connector));
                connectors.Add(connector);

                connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
                connector.Line.Weight = 1.5f;
                connector.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            }

            // Create bidirectional arrow between KG and ML
            PowerPointShape kgToMlConnector = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight,
                kgCircle.Left,
                centerY,
                mlCircle.Left + mlCircle.Width,
                centerY
            );
            ComReleaser.TrackObject(kgToMlConnector);
            Marshal.AddRef(Marshal.GetIUnknownForObject(kgToMlConnector));

            kgToMlConnector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            kgToMlConnector.Line.Weight = 2.5f;
            kgToMlConnector.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            kgToMlConnector.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;

            // Add labels for bidirectional arrows
            PowerPointShape kgToMlLabel = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius + 30,
                centerY - 50,
                180,
                40
            );
            ComReleaser.TrackObject(kgToMlLabel);
            Marshal.AddRef(Marshal.GetIUnknownForObject(kgToMlLabel));

            kgToMlLabel.TextFrame.TextRange.Text = "ML builds & enhances KGs";
            kgToMlLabel.TextFrame.TextRange.Font.Size = 11;
            kgToMlLabel.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            kgToMlLabel.Line.Visible = MsoTriState.msoFalse;

            PowerPointShape mlToKgLabel = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius + 10,
                centerY + 20,
                200,
                40
            );
            ComReleaser.TrackObject(mlToKgLabel);
            Marshal.AddRef(Marshal.GetIUnknownForObject(mlToKgLabel));

            mlToKgLabel.TextFrame.TextRange.Text = "KGs provide structured knowledge to ML";
            mlToKgLabel.TextFrame.TextRange.Font.Size = 11;
            mlToKgLabel.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            mlToKgLabel.Line.Visible = MsoTriState.msoFalse;

            // Add explanatory text box
            PowerPointShape explanation = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width - 250,
                slide.Design.SlideMaster.Height - 150,
                200,
                100
            );
            ComReleaser.TrackObject(explanation);
            Marshal.AddRef(Marshal.GetIUnknownForObject(explanation));

            explanation.TextFrame.TextRange.Text =
                "Knowledge graphs and machine learning have a symbiotic relationship. ML techniques can help build and enhance KGs, while KGs provide structured knowledge that improves ML models.";
            explanation.TextFrame.TextRange.Font.Size = 11;
            explanation.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            explanation.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            explanation.Line.Visible = MsoTriState.msoTrue;
            explanation.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(250, 250, 250));

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                try
                {
                    // Safely access notes page
                    if (slide.NotesPage != null && slide.NotesPage.Shapes.Count >= 2)
                    {
                        slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
                    }
                    else
                    {
                        Console.WriteLine("Warning: Cannot add notes - notes shape not available");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not add notes to the slide: {ex.Message}");
                }
            }

            // Animate the diagram (optional)
            if (slide.TimeLine != null)
            {
                try
                {
                    // Limit animations to avoid PowerPoint limits
                    List<PowerPointShape> allShapes = new List<PowerPointShape>
                    {
                        kgCircle,
                        kgText,
                        mlCircle,
                        mlText,
                        kgToMlConnector,
                        explanation
                    };
                    
                    // Add integration circles
                    if (integrationCircles != null)
                    {
                        // Limit to avoid exceeding PowerPoint's limit
                        int maxCirclesToAdd = Math.Min(integrationCircles.Length, 19); // 25 total - 6 already added = 19 max
                        for (int i = 0; i < maxCirclesToAdd; i++)
                        {
                            allShapes.Add(integrationCircles[i]);
                        }
                    }
                    
                    AnimateMLDiagram(slide, kgCircle, kgText, mlCircle, mlText, integrationCircles, kgToMlConnector, explanation);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not animate ML integration diagram: {ex.Message}");
                }
            }

            // Keep all shape lists alive until after animations
            GC.KeepAlive(integrationCircles);
            GC.KeepAlive(integrationTexts);
            GC.KeepAlive(connectors);

            return slide;
        }

        #region Helper Methods

        /// <summary>
        /// Adds a legend item to the slide
        /// </summary>
        /// <param name="slide">The slide to add the legend item to</param>
        /// <param name="x">X-coordinate</param>
        /// <param name="y">Y-coordinate</param>
        /// <param name="color">Color for the legend item</param>
        /// <param name="text">Text description</param>
        private void AddLegendItem(Slide slide, float x, float y, Color color, string text)
        {
            // Create color box
            PowerPointShape colorBox = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                x, y, 15, 15);
            ComReleaser.TrackObject(colorBox);
            
            colorBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
            colorBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);

            // Create label
            PowerPointShape label = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                x + 20, y - 2, 140, 20);
            ComReleaser.TrackObject(label);
            
            label.TextFrame.TextRange.Text = text;
            label.TextFrame.TextRange.Font.Size = 12;
            label.Line.Visible = MsoTriState.msoFalse;
        }

        /// <summary>
        /// Creates an entity node in the knowledge graph
        /// </summary>
        /// <param name="slide">The slide to add the node to</param>
        /// <param name="label">Label for the node</param>
        /// <param name="x">X-coordinate (center)</param>
        /// <param name="y">Y-coordinate (center)</param>
        /// <param name="color">Color for the node</param>
        /// <returns>The created shape</returns>
        private PowerPointShape CreateEntityNode(Slide slide, string label, float x, float y, Color color)
        {
            float size = 80;

            // Create entity node (circle)
            PowerPointShape node = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                x - size / 2, y - size / 2, size, size);
            ComReleaser.TrackObject(node);
            
            node.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
            node.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkGray);
            node.Line.Weight = 1.5f;

            // Make the node slightly 3D
            node.ThreeD.Visible = MsoTriState.msoTrue;
            node.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
            node.ThreeD.BevelTopInset = 3;
            node.ThreeD.BevelTopDepth = 3;

            // Add label
            node.TextFrame.TextRange.Text = label;
            node.TextFrame.TextRange.Font.Size = 14;
            node.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            node.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            node.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            node.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

            return node;
        }

        /// <summary>
        /// Adds a property badge to an entity node
        /// </summary>
        /// <param name="slide">The slide to add the property to</param>
        /// <param name="nodeShape">The node shape to attach the property to</param>
        /// <param name="propertyText">Text for the property</param>
        /// <param name="angle">Angle in degrees for positioning</param>
        /// <param name="color">Color for the property badge</param>
        private void AddPropertyBadge(Slide slide, PowerPointShape nodeShape, string propertyText, float angle, Color color)
        {
            // Calculate position based on angle
            float nodeX = nodeShape.Left + nodeShape.Width / 2;
            float nodeY = nodeShape.Top + nodeShape.Height / 2;
            float radius = nodeShape.Width / 2 + 30; // 30px away from node edge

            float radians = (float)(angle * Math.PI / 180.0);
            float propX = nodeX + (float)(radius * Math.Cos(radians));
            float propY = nodeY + (float)(radius * Math.Sin(radians));

            // Create property badge
            PowerPointShape propShape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                propX - 60, propY - 15, 120, 30);
            ComReleaser.TrackObject(propShape);
            
            propShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
            propShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(84, 130, 53));

            // Add text
            propShape.TextFrame.TextRange.Text = propertyText;
            propShape.TextFrame.TextRange.Font.Size = 10;
            propShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            propShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            propShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            propShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;

            // Create connector line from node to property
            PowerPointShape connector = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight,
                nodeX, nodeY, propX, propY);
            ComReleaser.TrackObject(connector);
            
            connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(84, 130, 53));
            connector.Line.Weight = 1.0f;
            connector.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;
        }

        /// <summary>
        /// Creates a relationship edge between two entity nodes
        /// </summary>
        /// <param name="slide">The slide to add the relationship to</param>
        /// <param name="startNode">Starting node</param>
        /// <param name="endNode">Ending node</param>
        /// <param name="label">Relationship label</param>
        /// <returns>The grouped shape containing the connector and label</returns>
        private PowerPointShape CreateRelationship(Slide slide, PowerPointShape startNode, PowerPointShape endNode, string label)
        {
            // Calculate connector points (center of nodes)
            float startX = startNode.Left + startNode.Width / 2;
            float startY = startNode.Top + startNode.Height / 2;
            float endX = endNode.Left + endNode.Width / 2;
            float endY = endNode.Top + endNode.Height / 2;

            // Create connector line
            PowerPointShape connector = slide.Shapes.AddLine(startX, startY, endX, endY);
            ComReleaser.TrackObject(connector);
            
            connector.Line.ForeColor.RGB = ColorTranslator.ToOle(edgeColor);
            connector.Line.Weight = 2.0f;

            // Add arrowhead
            connector.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            connector.Line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadLengthMedium;
            connector.Line.EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWidthMedium;

            // Calculate midpoint for label
            float midX = (startX + endX) / 2;
            float midY = (startY + endY) / 2;

            // Add relationship label background
            PowerPointShape labelBg = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                midX - 45, midY - 12, 90, 24);
            ComReleaser.TrackObject(labelBg);
            
            labelBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            labelBg.Fill.Transparency = 0.2f;
            labelBg.Line.ForeColor.RGB = ColorTranslator.ToOle(edgeColor);
            labelBg.Line.Weight = 1.0f;

            // Add relationship label
            PowerPointShape labelShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                midX - 40, midY - 10, 80, 20);
            ComReleaser.TrackObject(labelShape);
            
            labelShape.TextFrame.TextRange.Text = label;
            labelShape.TextFrame.TextRange.Font.Size = 10;
            labelShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            labelShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(edgeColor);
            labelShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            labelShape.Line.Visible = MsoTriState.msoFalse;

            // Group connector and label for animation purposes
            int[] shapeIds = new int[] { connector.Id, labelBg.Id, labelShape.Id };
            PowerPointShapeRange shapeRange = slide.Shapes.Range(shapeIds);
            ComReleaser.TrackObject(shapeRange);
            
            PowerPointShape groupedShape = shapeRange.Group();
            ComReleaser.TrackObject(groupedShape);

            return groupedShape;
        }

        /// <summary>
        /// Mapping a node index to an animation effect
        /// </summary>
        /// <param name="index">Index of the node</param>
        /// <returns>Animation effect for the node</returns>
        private MsoAnimEffect GetNodeAnimationEffect(int index)
        {
            return MsoAnimEffect.msoAnimEffectFade;
        }

        /// <summary>
        /// Adds animations to the ML diagram elements
        /// </summary>
        private void AnimateMLDiagram(
            Slide slide,
            PowerPointShape kgCircle,
            PowerPointShape kgText,
            PowerPointShape mlCircle,
            PowerPointShape mlText,
            PowerPointShape[] integrationCircles,
            PowerPointShape kgToMlConnector,
            PowerPointShape explanation)
        {
            try
            {
                // Check if TimeLine and MainSequence are accessible
                if (slide.TimeLine == null || slide.TimeLine.MainSequence == null)
                {
                    Console.WriteLine("Cannot add animations: Timeline or MainSequence is null");
                    return;
                }
                
                // Calculate available animation slots
                int maxAnimations = 25; // PowerPoint limit is around 25-30 animations
                int currentCount = 0;
                int availableSlots = maxAnimations;
                
                // Track how many animations we've added
                Console.WriteLine("Starting ML diagram animation sequence");
                
                // First add essential animations
                if (currentCount < availableSlots && kgCircle != null)
                {
                    Effect kgCircleEffect = slide.TimeLine.MainSequence.AddEffect(
                        kgCircle,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    ComReleaser.TrackObject(kgCircleEffect);
                    kgCircleEffect.Timing.Duration = 0.5f;
                    currentCount++;
                }
                
                if (currentCount < availableSlots && kgText != null)
                {
                    Effect kgTextEffect = slide.TimeLine.MainSequence.AddEffect(
                        kgText,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    ComReleaser.TrackObject(kgTextEffect);
                    kgTextEffect.Timing.Duration = 0.5f;
                    currentCount++;
                }
                
                // Add integration circles with limit
                if (integrationCircles != null)
                {
                    int circleCount = Math.Min(integrationCircles.Length, availableSlots - currentCount);
                    Console.WriteLine($"Animating {circleCount} integration circles");
                    
                    for (int i = 0; i < circleCount; i++)
                    {
                        if (integrationCircles[i] != null)
                        {
                            Effect circleEffect = slide.TimeLine.MainSequence.AddEffect(
                                integrationCircles[i],
                                MsoAnimEffect.msoAnimEffectFade,
                                MsoAnimateByLevel.msoAnimateLevelNone,
                                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            ComReleaser.TrackObject(circleEffect);
                            circleEffect.Timing.Duration = 0.3f;
                            currentCount++;
                            
                            // Release older COM objects periodically
                            if (i % 3 == 0)
                            {
                                ComReleaser.ReleaseOldestObjects(5);
                            }
                        }
                    }
                }
                
                // Add ML circle and text if slots available 
                if (currentCount < availableSlots && mlCircle != null)
                {
                    Effect mlCircleEffect = slide.TimeLine.MainSequence.AddEffect(
                        mlCircle,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(mlCircleEffect);
                    mlCircleEffect.Timing.Duration = 0.5f;
                    currentCount++;
                }
                
                if (currentCount < availableSlots && mlText != null)
                {
                    Effect mlTextEffect = slide.TimeLine.MainSequence.AddEffect(
                        mlText,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    ComReleaser.TrackObject(mlTextEffect);
                    mlTextEffect.Timing.Duration = 0.5f;
                    currentCount++;
                }
                
                // Add connector if slots available
                if (currentCount < availableSlots && kgToMlConnector != null)
                {
                    Effect connectorEffect = slide.TimeLine.MainSequence.AddEffect(
                        kgToMlConnector,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(connectorEffect);
                    connectorEffect.Timing.Duration = 0.5f;
                    currentCount++;
                }
                
                // Add explanation text if slots available 
                if (currentCount < availableSlots && explanation != null)
                {
                    Effect explanationEffect = slide.TimeLine.MainSequence.AddEffect(
                        explanation,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(explanationEffect);
                    explanationEffect.Timing.Duration = 0.7f;
                    currentCount++;
                }
                
                Console.WriteLine($"Added {currentCount} animations to ML diagram");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error animating ML diagram: {ex.Message}");
            }
        }

        #endregion
    }
}