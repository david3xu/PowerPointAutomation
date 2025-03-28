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
    /// Class for generating list-based slides with bulleted or numbered content
    /// </summary>
    public class ListSlide
    {
        // Constants for missing enum values
        private const PpPlaceholderType ppPlaceholderContent = (PpPlaceholderType)2;
        private const MsoAnimateByLevel msoAnimateLevelParagraphs = (MsoAnimateByLevel)2;
        private const MsoAnimDirection msoAnimDirectionFromLeft = (MsoAnimDirection)3;
        
        private Presentation presentation;
        private CustomLayout slideLayout;
        
        // Brand colors
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        /// <summary>
        /// Initializes a new instance of the ListSlide class
        /// </summary>
        /// <param name="presentation">PowerPoint presentation</param>
        /// <param name="layout">Slide layout to use</param>
        public ListSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.slideLayout = layout;
        }

        /// <summary>
        /// Generates a slide with a bulleted list
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="items">Array of bullet point texts</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <returns>The created slide</returns>
        public Slide GenerateBulletedList(string title, string[] items, string notes = null)
        {
            Slide slide = null;
            try
            {
                // Add slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, slideLayout);
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

                // Create content background
                PowerPointShape contentBg = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    50, 100, slide.Design.SlideMaster.Width - 100, slide.Design.SlideMaster.Height - 150);
                ComReleaser.TrackObject(contentBg);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(contentBg));
                
                contentBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
                contentBg.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                contentBg.Line.Weight = 1.0f;
                contentBg.ZOrder(MsoZOrderCmd.msoSendToBack);

                // Calculate item layout
                float itemHeight = 40;
                float spacing = 10;
                float totalItemsHeight = items.Length * itemHeight + (items.Length - 1) * spacing;
                float startY = 100 + (slide.Design.SlideMaster.Height - 250 - totalItemsHeight) / 2;

                // Keep track of all shapes for animations
                List<PowerPointShape> itemShapes = new List<PowerPointShape>();

                // Add bullet points
                for (int i = 0; i < items.Length; i++)
                {
                    float itemY = startY + i * (itemHeight + spacing);
                    
                    // Create bullet point background
                    PowerPointShape itemBg = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        70, itemY, slide.Design.SlideMaster.Width - 140, itemHeight);
                    ComReleaser.TrackObject(itemBg);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemBg));
                    
                    itemShapes.Add(itemBg);
                    
                    itemBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
                    itemBg.Fill.Transparency = 0.7f;
                    itemBg.Line.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
                    itemBg.Line.Weight = 1.0f;

                    // Add bullet point text
                    PowerPointShape itemText = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        90, itemY + 5, slide.Design.SlideMaster.Width - 160, itemHeight - 10);
                    ComReleaser.TrackObject(itemText);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(itemText));
                    
                    itemText.TextFrame.TextRange.Text = items[i];
                    itemText.TextFrame.TextRange.Font.Size = 18;
                    itemText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51)); // Dark gray
                    itemText.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    itemText.Line.Visible = MsoTriState.msoFalse;
                }

                // Keep all shape lists alive until after animations
                GC.KeepAlive(itemShapes);

                // Add animations
                if (slide.TimeLine != null)
                {
                    // Animate title
                    Effect titleEffect = slide.TimeLine.MainSequence.AddEffect(
                        titleShape,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    ComReleaser.TrackObject(titleEffect);

                    // Animate background
                    Effect bgEffect = slide.TimeLine.MainSequence.AddEffect(
                        contentBg,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(bgEffect);

                    // Animate bullet points
                    for (int i = 0; i < itemShapes.Count; i++)
                    {
                        Effect itemEffect = slide.TimeLine.MainSequence.AddEffect(
                            itemShapes[i],
                            MsoAnimEffect.msoAnimEffectFly,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            i == 0 ? MsoAnimTriggerType.msoAnimTriggerAfterPrevious : MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        ComReleaser.TrackObject(itemEffect);
                        
                        itemEffect.EffectParameters.Direction = msoAnimDirectionFromLeft;
                        itemEffect.Timing.Duration = 0.5f;
                        
                        if (i > 0)
                        {
                            itemEffect.Timing.TriggerDelayTime = 0.2f;
                        }
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
                Console.WriteLine($"Error generating bulleted list slide: {ex.Message}");
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