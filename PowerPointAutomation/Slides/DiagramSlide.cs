using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

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
        /// <param name="animate">Whether to animate the diagram</param>
        /// <returns>The created slide</returns>
        public Slide GenerateKnowledgeGraphDiagram(string title, string subtitle = null, string notes = null, bool animate = true)
        {
            // Add diagram slide
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Set title
            slide.Shapes.Title.TextFrame.TextRange.Text = title;
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add subtitle if provided
            if (!string.IsNullOrEmpty(subtitle))
            {
                float subtitleTop = slide.Shapes.Title.Top + slide.Shapes.Title.Height + 10;
                Shape subtitleShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    slide.Shapes.Title.Left,
                    subtitleTop,
                    slide.Design.SlideMaster.Width - (slide.Shapes.Title.Left * 2),
                    30
                );

                subtitleShape.TextFrame.TextRange.Text = subtitle;
                subtitleShape.TextFrame.TextRange.Font.Size = 18;
                subtitleShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
                subtitleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                subtitleShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                subtitleShape.Line.Visible = MsoTriState.msoFalse;
                subtitleShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12;
            }

            // Create background for the diagram
            float diagramLeft = 100;
            float diagramTop = 150;
            float diagramWidth = slide.Design.SlideMaster.Width - 200;
            float diagramHeight = slide.Design.SlideMaster.Height - 250;

            Shape diagramBackground = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                diagramLeft,
                diagramTop,
                diagramWidth,
                diagramHeight
            );

            // Format diagram background
            diagramBackground.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
            diagramBackground.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            diagramBackground.Line.Weight = 1.0f;

            // Create legend
            Shape legendBox = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                diagramLeft + 20,
                diagramTop + 20,
                180,
                120
            );

            legendBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            legendBox.Fill.Transparency = 0.2f;
            legendBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
            legendBox.Line.Weight = 1.0f;

            Shape legendTitle = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                diagramLeft + 20,
                diagramTop + 20,
                180,
                30
            );

            legendTitle.TextFrame.TextRange.Text = "Legend";
            legendTitle.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            legendTitle.TextFrame.TextRange.Font.Size = 14;
            legendTitle.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            legendTitle.Line.Visible = MsoTriState.msoFalse;

            // Add legend items
            AddLegendItem(slide, diagramLeft + 30, diagramTop + 55, Color.FromArgb(91, 155, 213), "Entity (Node)");
            AddLegendItem(slide, diagramLeft + 30, diagramTop + 85, Color.FromArgb(237, 125, 49), "Relationship (Edge)");
            AddLegendItem(slide, diagramLeft + 30, diagramTop + 115, Color.FromArgb(112, 173, 71), "Property");

            // Create nodes (entities) in the knowledge graph
            // Calculate positions for a better layout
            float centerX = diagramLeft + diagramWidth / 2;
            float centerY = diagramTop + diagramHeight / 2;
            float radius = Math.Min(diagramWidth, diagramHeight) * 0.3f;

            Shape companyNode = CreateEntityNode(slide, "Company", centerX, centerY - radius, nodeColor);
            Shape personNode = CreateEntityNode(slide, "Person", centerX - radius, centerY, nodeColor);
            Shape productNode = CreateEntityNode(slide, "Product", centerX + radius * 0.7f, centerY + radius * 0.7f, nodeColor);
            Shape featureNode = CreateEntityNode(slide, "Feature", centerX + radius, centerY, nodeColor);

            // Add properties to company node
            AddPropertyBadge(slide, companyNode, "name: 'TechCorp'", -45, Color.FromArgb(112, 173, 71));
            AddPropertyBadge(slide, companyNode, "founded: 2010", 0, Color.FromArgb(112, 173, 71));

            // Add property to person node
            AddPropertyBadge(slide, personNode, "name: 'John'", 45, Color.FromArgb(112, 173, 71));

            // Add property to product node
            AddPropertyBadge(slide, productNode, "category: 'Software'", 0, Color.FromArgb(112, 173, 71));

            // Create edges (relationships) in the knowledge graph
            Shape employeesEdge = CreateRelationship(slide, companyNode, personNode, "EMPLOYS");
            Shape producesEdge = CreateRelationship(slide, companyNode, productNode, "PRODUCES");
            Shape featuresEdge = CreateRelationship(slide, productNode, featureNode, "HAS_FEATURE");
            Shape developsEdge = CreateRelationship(slide, personNode, productNode, "DEVELOPS");

            // Add animation if requested
            if (animate)
            {
                // Group shapes for animation
                Shape[] nodeShapes = new Shape[] { companyNode, personNode, productNode, featureNode };
                Shape[] edgeShapes = new Shape[] { employeesEdge, producesEdge, featuresEdge, developsEdge };

                // First animate the background and legend
                Effect bgEffect = slide.TimeLine.MainSequence.AddEffect(
                    diagramBackground,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerOnClick);

                Effect legendEffect = slide.TimeLine.MainSequence.AddEffect(
                    legendBox,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                Effect legendTitleEffect = slide.TimeLine.MainSequence.AddEffect(
                    legendTitle,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                // Animate nodes one by one
                for (int i = 0; i < nodeShapes.Length; i++)
                {
                    Effect nodeEffect = slide.TimeLine.MainSequence.AddEffect(
                        nodeShapes[i],
                        MsoAnimEffect.msoAnimEffectFly,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        i == 0 ? MsoAnimTriggerType.msoAnimTriggerAfterPrevious : MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    nodeEffect.EffectParameters.Direction = GetNodeAnimationDirection(i);
                    nodeEffect.Timing.Duration = 0.5f;
                }

                // Animate edges after nodes
                for (int i = 0; i < edgeShapes.Length; i++)
                {
                    Effect edgeEffect = slide.TimeLine.MainSequence.AddEffect(
                        edgeShapes[i],
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    edgeEffect.Timing.Duration = 0.4f;
                }
            }

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

            return slide;
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

            // Set title
            slide.Shapes.Title.TextFrame.TextRange.Text = title;
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add subtitle if provided
            if (!string.IsNullOrEmpty(subtitle))
            {
                float subtitleTop = slide.Shapes.Title.Top + slide.Shapes.Title.Height + 10;
                Shape subtitleShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    slide.Shapes.Title.Left,
                    subtitleTop,
                    slide.Design.SlideMaster.Width - (slide.Shapes.Title.Left * 2),
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
            Shape kgCircle = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                centerX - 80,
                centerY - 80,
                160,
                160
            );

            kgCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196)); // Blue
            kgCircle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125));
            kgCircle.Line.Weight = 2.0f;

            Shape kgText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - 70,
                centerY - 25,
                140,
                50
            );

            kgText.TextFrame.TextRange.Text = "Knowledge Graphs";
            kgText.TextFrame.TextRange.Font.Size = 16;
            kgText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            kgText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            kgText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            kgText.Line.Visible = MsoTriState.msoFalse;

            // Create ML circle
            Shape mlCircle = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                centerX - radius - 60,
                centerY - 60,
                120,
                120
            );

            mlCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(112, 173, 71)); // Green
            mlCircle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(84, 130, 53));
            mlCircle.Line.Weight = 2.0f;

            Shape mlText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius - 50,
                centerY - 20,
                100,
                50
            );

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

            Shape[] integrationCircles = new Shape[integrationPoints.Length];

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

                integrationCircles[i].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(237, 125, 49)); // Orange
                integrationCircles[i].Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(191, 96, 27));
                integrationCircles[i].Line.Weight = 1.5f;

                Shape integrationText = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    x - 45,
                    y - 25,
                    90,
                    50
                );

                integrationText.TextFrame.TextRange.Text = integrationPoints[i];
                integrationText.TextFrame.TextRange.Font.Size = 11;
                integrationText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                integrationText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                integrationText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                integrationText.Line.Visible = MsoTriState.msoFalse;

                // Create connector line from KG to integration point
                Shape connector = slide.Shapes.AddConnector(
                    MsoConnectorType.msoConnectorStraight,
                    centerX,
                    centerY,
                    x,
                    y
                );

                connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
                connector.Line.Weight = 1.5f;
                connector.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            }

            // Create bidirectional arrow between KG and ML
            Shape kgToMlConnector = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight,
                kgCircle.Left,
                centerY,
                mlCircle.Left + mlCircle.Width,
                centerY
            );

            kgToMlConnector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            kgToMlConnector.Line.Weight = 2.5f;
            kgToMlConnector.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
            kgToMlConnector.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;

            // Add labels for bidirectional arrows
            Shape kgToMlLabel = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius + 30,
                centerY - 50,
                180,
                40
            );

            kgToMlLabel.TextFrame.TextRange.Text = "ML builds & enhances KGs";
            kgToMlLabel.TextFrame.TextRange.Font.Size = 11;
            kgToMlLabel.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            kgToMlLabel.Line.Visible = MsoTriState.msoFalse;

            Shape mlToKgLabel = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                centerX - radius + 10,
                centerY + 20,
                200,
                40
            );

            mlToKgLabel.TextFrame.TextRange.Text = "KGs provide structured knowledge to ML";
            mlToKgLabel.TextFrame.TextRange.Font.Size = 11;
            mlToKgLabel.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            mlToKgLabel.Line.Visible = MsoTriState.msoFalse;

            // Add explanatory text box
            Shape explanation = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width - 250,
                slide.Design.SlideMaster.Height - 150,
                200,
                100
            );

            explanation.TextFrame.TextRange.Text =
                "Knowledge graphs and machine learning have a symbiotic relationship. ML techniques can help build and enhance KGs, while KGs provide structured knowledge that improves ML models.";
            explanation.TextFrame.TextRange.Font.Size = 11;
            explanation.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            explanation.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            explanation.Line.Visible = MsoTriState.msoTrue;
            explanation.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(250, 250, 250));

            // Add animations to diagram elements
            AnimateMLDiagram(slide, kgCircle, kgText, mlCircle, mlText, integrationCircles, kgToMlConnector, explanation);

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

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
            Shape colorBox = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                x, y, 15, 15);
            colorBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
            colorBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);

            // Create label
            Shape label = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                x + 20, y - 2, 140, 20);
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
        private Shape CreateEntityNode(Slide slide, string label, float x, float y, Color color)
        {
            float size = 80;

            // Create entity node (circle)
            Shape node = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                x - size / 2, y - size / 2, size, size);
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
        private void AddPropertyBadge(Slide slide, Shape nodeShape, string propertyText, float angle, Color color)
        {
            // Calculate position based on angle
            float nodeX = nodeShape.Left + nodeShape.Width / 2;
            float nodeY = nodeShape.Top + nodeShape.Height / 2;
            float radius = nodeShape.Width / 2 + 30; // 30px away from node edge

            float radians = (float)(angle * Math.PI / 180.0);
            float propX = nodeX + (float)(radius * Math.Cos(radians));
            float propY = nodeY + (float)(radius * Math.Sin(radians));

            // Create property badge
            Shape propShape = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                propX - 60, propY - 15, 120, 30);
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
            Shape connector = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight,
                nodeX, nodeY, propX, propY);
            connector.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(84, 130, 53));
            connector.Line.Weight = 1.0f;
            connector.Line.DashStyle = MsoLineDashStyle.msoDashDot;
        }

        /// <summary>
        /// Creates a relationship edge between two entity nodes
        /// </summary>
        /// <param name="slide">The slide to add the relationship to</param>
        /// <param name="startNode">Starting node</param>
        /// <param name="endNode">Ending node</param>
        /// <param name="label">Relationship label</param>
        /// <returns>The grouped shape containing the connector and label</returns>
        private Shape CreateRelationship(Slide slide, Shape startNode, Shape endNode, string label)
        {
            // Calculate connector points (center of nodes)
            float startX = startNode.Left + startNode.Width / 2;
            float startY = startNode.Top + startNode.Height / 2;
            float endX = endNode.Left + endNode.Width / 2;
            float endY = endNode.Top + endNode.Height / 2;

            // Create connector line
            Shape connector = slide.Shapes.AddLine(startX, startY, endX, endY);
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
            Shape labelBg = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                midX - 45, midY - 12, 90, 24);
            labelBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
            labelBg.Fill.Transparency = 0.2f;
            labelBg.Line.ForeColor.RGB = ColorTranslator.ToOle(edgeColor);
            labelBg.Line.Weight = 1.0f;

            // Add relationship label
            Shape labelShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                midX - 40, midY - 10, 80, 20);
            labelShape.TextFrame.TextRange.Text = label;
            labelShape.TextFrame.TextRange.Font.Size = 10;
            labelShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            labelShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(edgeColor);
            labelShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            labelShape.Line.Visible = MsoTriState.msoFalse;

            // Group connector and label for animation purposes
            ShapeRange shapeRange = slide.Shapes.Range(new int[] { connector.Id, labelBg.Id, labelShape.Id });
            Shape groupedShape = shapeRange.Group();

            return groupedShape;
        }

        /// <summary>
        /// Gets the animation direction for a node based on its index
        /// </summary>
        /// <param name="index">Index of the node</param>
        /// <returns>Animation direction</returns>
        private MsoAnimDirection GetNodeAnimationDirection(int index)
        {
            switch (index % 4)
            {
                case 0:
                    return MsoAnimDirection.msoAnimDirectionFromTop;
                case 1:
                    return MsoAnimDirection.msoAnimDirectionFromLeft;
                case 2:
                    return MsoAnimDirection.msoAnimDirectionFromRight;
                case 3:
                    return MsoAnimDirection.msoAnimDirectionFromBottom;
                default:
                    return MsoAnimDirection.msoAnimDirectionFromTop;
            }
        }

        /// <summary>
        /// Adds animations to the ML diagram elements
        /// </summary>
        private void AnimateMLDiagram(
            Slide slide,
            Shape kgCircle,
            Shape kgText,
            Shape mlCircle,
            Shape mlText,
            Shape[] integrationCircles,
            Shape kgToMlConnector,
            Shape explanation)
        {
            // First animate KG circle and text
            Effect kgCircleEffect = slide.TimeLine.MainSequence.AddEffect(
                kgCircle,
                MsoAnimEffect.msoAnimEffectGrowAndTurn,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerOnClick);

            Effect kgTextEffect = slide.TimeLine.MainSequence.AddEffect(
                kgText,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            // Then animate ML circle and text
            Effect mlCircleEffect = slide.TimeLine.MainSequence.AddEffect(
                mlCircle,
                MsoAnimEffect.msoAnimEffectGrowAndTurn,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            Effect mlTextEffect = slide.TimeLine.MainSequence.AddEffect(
                mlText,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            // Animate connector between KG and ML
            Effect connectorEffect = slide.TimeLine.MainSequence.AddEffect(
                kgToMlConnector,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            // Animate integration circles one by one
            for (int i = 0; i < integrationCircles.Length; i++)
            {
                Effect circleEffect = slide.TimeLine.MainSequence.AddEffect(
                    integrationCircles[i],
                    MsoAnimEffect.msoAnimEffectFly,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                circleEffect.EffectParameters.Direction = GetNodeAnimationDirection(i);
            }

            // Finally, fade in the explanation
            Effect explanationEffect = slide.TimeLine.MainSequence.AddEffect(
                explanation,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
        }

        #endregion
    }
}