using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointAutomation.Utilities;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class to create a slide with two columns of content for comparisons
    /// </summary>
    public class ComparisonSlide
    {
        private Presentation presentation;
        private CustomLayout layout;
        
        // Brand colors
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange
        private readonly Color leftColor = Color.FromArgb(91, 155, 213);      // Light blue
        private readonly Color rightColor = Color.FromArgb(112, 173, 71);     // Green

        public ComparisonSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a slide with two columns of content
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="leftItems">Array of items for the left column</param>
        /// <param name="rightItems">Array of items for the right column</param>
        /// <param name="leftHeader">Optional header for the left column</param>
        /// <param name="rightHeader">Optional header for the right column</param>
        /// <param name="notes">Optional slide notes</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string[] leftItems, string[] rightItems, 
            string leftHeader = null, string rightHeader = null, string notes = null)
        {
            Slide slide = null;
            try
            {
                // Add slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);
                ComReleaser.TrackObject(slide);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(slide));

                // Try to get the title shape
                PowerPointShape titleShape = GetOrCreateTitleShape(slide);
                ComReleaser.TrackObject(titleShape);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(titleShape));
                
                titleShape.TextFrame.TextRange.Text = title;
                titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                titleShape.TextFrame.TextRange.Font.Size = 36;
                titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

                // Calculate layout dimensions
                float contentTop = 100;
                float contentHeight = slide.Design.SlideMaster.Height - 150;
                float columnWidth = (slide.Design.SlideMaster.Width - 150) / 2;
                float leftColumnLeft = 50;
                float rightColumnLeft = leftColumnLeft + columnWidth + 50;
                
                // Create left column header if provided
                PowerPointShape leftHeaderShape = null;
                PowerPointShape leftHeaderBg = null;
                if (!string.IsNullOrEmpty(leftHeader))
                {
                    leftHeaderShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        leftColumnLeft,
                        contentTop,
                        columnWidth,
                        40);
                    ComReleaser.TrackObject(leftHeaderShape);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(leftHeaderShape));
                    
                    leftHeaderShape.TextFrame.TextRange.Text = leftHeader;
                    leftHeaderShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    leftHeaderShape.TextFrame.TextRange.Font.Size = 24;
                    leftHeaderShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(leftColor);
                    leftHeaderShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    leftHeaderShape.Line.Visible = MsoTriState.msoFalse;
                    
                    // Add background for header
                    leftHeaderBg = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        leftColumnLeft,
                        contentTop,
                        columnWidth,
                        40);
                    ComReleaser.TrackObject(leftHeaderBg);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(leftHeaderBg));
                    
                    leftHeaderBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(235, 241, 249)); // Light blue
                    leftHeaderBg.Line.ForeColor.RGB = ColorTranslator.ToOle(leftColor);
                    leftHeaderBg.Line.Weight = 1.5f;
                    leftHeaderBg.ZOrder(MsoZOrderCmd.msoSendBackward);
                    
                    // Adjust content top position
                    contentTop += 50;
                    contentHeight -= 50;
                }
                
                // Create right column header if provided
                PowerPointShape rightHeaderShape = null;
                PowerPointShape rightHeaderBg = null;
                if (!string.IsNullOrEmpty(rightHeader))
                {
                    rightHeaderShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        rightColumnLeft,
                        contentTop,
                        columnWidth,
                        40);
                    ComReleaser.TrackObject(rightHeaderShape);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(rightHeaderShape));
                    
                    rightHeaderShape.TextFrame.TextRange.Text = rightHeader;
                    rightHeaderShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    rightHeaderShape.TextFrame.TextRange.Font.Size = 24;
                    rightHeaderShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(rightColor);
                    rightHeaderShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    rightHeaderShape.Line.Visible = MsoTriState.msoFalse;
                    
                    // Add background for header
                    rightHeaderBg = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        rightColumnLeft,
                        contentTop,
                        columnWidth,
                        40);
                    ComReleaser.TrackObject(rightHeaderBg);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(rightHeaderBg));
                    
                    rightHeaderBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(235, 241, 249)); // Light blue
                    rightHeaderBg.Line.ForeColor.RGB = ColorTranslator.ToOle(rightColor);
                    rightHeaderBg.Line.Weight = 1.5f;
                    rightHeaderBg.ZOrder(MsoZOrderCmd.msoSendBackward);
                    
                    // Adjust content top position
                    contentTop += 50;
                    contentHeight -= 50;
                }

                // Create left column background
                PowerPointShape leftColumnBg = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    leftColumnLeft,
                    contentTop,
                    columnWidth,
                    contentHeight);
                ComReleaser.TrackObject(leftColumnBg);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(leftColumnBg));
                
                leftColumnBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
                leftColumnBg.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                leftColumnBg.Line.Weight = 1.0f;
                leftColumnBg.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                leftColumnBg.ZOrder(MsoZOrderCmd.msoSendToBack);

                // Create right column background
                PowerPointShape rightColumnBg = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    rightColumnLeft,
                    contentTop,
                    columnWidth,
                    contentHeight);
                ComReleaser.TrackObject(rightColumnBg);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(rightColumnBg));
                
                rightColumnBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
                rightColumnBg.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                rightColumnBg.Line.Weight = 1.0f;
                rightColumnBg.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                rightColumnBg.ZOrder(MsoZOrderCmd.msoSendToBack);
                
                // Add vertical divider
                PowerPointShape divider = slide.Shapes.AddLine(
                    leftColumnLeft + columnWidth + 25, contentTop,
                    leftColumnLeft + columnWidth + 25, contentTop + contentHeight);
                ComReleaser.TrackObject(divider);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(divider));
                
                divider.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                divider.Line.Weight = 2.0f;
                divider.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                
                // Calculate item layout
                int maxItems = Math.Max(leftItems.Length, rightItems.Length);
                float itemHeight = 40;
                float spacing = 10;
                float totalItemsHeight = maxItems * itemHeight + (maxItems - 1) * spacing;
                float startY = contentTop + (contentHeight - totalItemsHeight) / 2;

                // Keep track of all shapes for animations
                List<PowerPointShape> leftItemShapes = new List<PowerPointShape>();
                List<PowerPointShape> rightItemShapes = new List<PowerPointShape>();

                // Add left column items
                for (int i = 0; i < leftItems.Length; i++)
                {
                    float itemY = startY + i * (itemHeight + spacing);
                    
                    PowerPointShape itemBg = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        leftColumnLeft + 20,
                        itemY,
                        columnWidth - 40,
                        itemHeight);
                    ComReleaser.TrackObject(itemBg);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemBg));
                    
                    leftItemShapes.Add(itemBg);
                    
                    itemBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(leftColor);
                    itemBg.Fill.Transparency = 0.7f;
                    itemBg.Line.ForeColor.RGB = ColorTranslator.ToOle(leftColor);
                    itemBg.Line.Weight = 1.0f;
                    
                    PowerPointShape itemText = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        leftColumnLeft + 25,
                        itemY + 5,
                        columnWidth - 50,
                        itemHeight - 10);
                    ComReleaser.TrackObject(itemText);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemText));
                    
                    itemText.TextFrame.TextRange.Text = leftItems[i];
                    itemText.TextFrame.TextRange.Font.Size = 18;
                    itemText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125)); // Dark blue
                    itemText.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    itemText.Line.Visible = MsoTriState.msoFalse;
                }

                // Add right column items
                for (int i = 0; i < rightItems.Length; i++)
                {
                    float itemY = startY + i * (itemHeight + spacing);
                    
                    PowerPointShape itemBg = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        rightColumnLeft + 20,
                        itemY,
                        columnWidth - 40,
                        itemHeight);
                    ComReleaser.TrackObject(itemBg);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemBg));
                    
                    rightItemShapes.Add(itemBg);
                    
                    itemBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(rightColor);
                    itemBg.Fill.Transparency = 0.7f;
                    itemBg.Line.ForeColor.RGB = ColorTranslator.ToOle(rightColor);
                    itemBg.Line.Weight = 1.0f;
                    
                    PowerPointShape itemText = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        rightColumnLeft + 25,
                        itemY + 5,
                        columnWidth - 50,
                        itemHeight - 10);
                    ComReleaser.TrackObject(itemText);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemText));
                    
                    itemText.TextFrame.TextRange.Text = rightItems[i];
                    itemText.TextFrame.TextRange.Font.Size = 18;
                    itemText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(21, 80, 21)); // Dark green
                    itemText.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    itemText.Line.Visible = MsoTriState.msoFalse;
                }

                // Keep all shape lists alive until after animations
                GC.KeepAlive(leftItemShapes);
                GC.KeepAlive(rightItemShapes);
                
                // Add animation effects
                if (slide.TimeLine != null)
                {
                    // First animate the title
                    Effect titleEffect = slide.TimeLine.MainSequence.AddEffect(
                        titleShape,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    ComReleaser.TrackObject(titleEffect);
                    
                    // Animate the divider
                    Effect dividerEffect = slide.TimeLine.MainSequence.AddEffect(
                        divider,
                        MsoAnimEffect.msoAnimEffectGrowShrink,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(dividerEffect);
                    
                    dividerEffect.Timing.Duration = 0.5f;
                    
                    // Animate left column items
                    for (int i = 0; i < leftItemShapes.Count; i++)
                    {
                        Effect effect = slide.TimeLine.MainSequence.AddEffect(
                            leftItemShapes[i],
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            i == 0 ? MsoAnimTriggerType.msoAnimTriggerAfterPrevious : MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        ComReleaser.TrackObject(effect);
                        
                        effect.Timing.Duration = 0.3f;
                    }
                    
                    // Animate right column items
                    for (int i = 0; i < rightItemShapes.Count; i++)
                    {
                        Effect effect = slide.TimeLine.MainSequence.AddEffect(
                            rightItemShapes[i],
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            i == 0 ? MsoAnimTriggerType.msoAnimTriggerAfterPrevious : MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        ComReleaser.TrackObject(effect);
                        
                        effect.Timing.Duration = 0.3f;
                    }
                }

                // Add speaker notes if provided
                if (!string.IsNullOrEmpty(notes))
                {
                    try
                    {
                        PowerPointShape notesShape = slide.NotesPage.Shapes[2];
                        ComReleaser.TrackObject(notesShape);
                        notesShape.TextFrame.TextRange.Text = notes;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not add notes to the slide: {ex.Message}");
                    }
                }

                return slide;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating comparison slide: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return slide; // Return the partially created slide or null
            }
        }
        
        /// <summary>
        /// Gets the title shape from the slide or creates a new one if it doesn't exist
        /// </summary>
        private PowerPointShape GetOrCreateTitleShape(Slide slide)
        {
            PowerPointShape titleShape = null;
            
            // Try to get the title placeholder
            try
            {
                foreach (PowerPointShape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder)
                    {
                        if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle)
                        {
                            titleShape = shape;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding title placeholder: {ex.Message}");
            }
            
            // If no title placeholder found, create a custom title shape
            if (titleShape == null)
            {
                titleShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    50, // Left
                    20, // Top
                    slide.Design.SlideMaster.Width - 100, // Width
                    50 // Height
                );
                
                // Format as title
                titleShape.TextFrame.TextRange.Font.Size = 36;
                titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                titleShape.Line.Visible = MsoTriState.msoFalse;
            }
            
            return titleShape;
        }
    }
} 