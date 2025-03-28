using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointAutomation.Utilities;
using System.Runtime.InteropServices;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class to create a summary or conclusion slide
    /// </summary>
    public class SummarySlide
    {
        // Constants for missing enum values
        private const MsoAnimDirection msoAnimDirectionFromLeft = (MsoAnimDirection)3;
        
        private Presentation presentation;
        private CustomLayout layout;
        
        // Brand colors
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        public SummarySlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a summary slide with title, content, and optional contact information
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="content">The main content text</param>
        /// <param name="contactInfo">Optional contact information</param>
        /// <param name="notes">Optional slide notes</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string content, string contactInfo = null, string notes = null)
        {
            Slide slide = null;
            try
            {
                // Add slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);
                ComReleaser.TrackObject(slide);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(slide));

                // Try to get the title shape from the placeholder
                PowerPointShape titleShape = null;
                try
                {
                    titleShape = GetOrCreateTitleShape(slide);
                    ComReleaser.TrackObject(titleShape);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(titleShape));
                    
                    // Set the title text
                    titleShape.TextFrame.TextRange.Text = title;
                    titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    titleShape.TextFrame.TextRange.Font.Size = 36;
                    titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not set title properly: {ex.Message}. Creating custom title.");
                    
                    // Create a custom title
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
                
                // Perform intermediate cleanup after title creation
                ComReleaser.ReleaseOldestObjects(5);

                // Create background accent
                PowerPointShape accentShape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    0, 0,
                    slide.Design.SlideMaster.Width, 5);
                ComReleaser.TrackObject(accentShape);
                
                accentShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                accentShape.Line.Visible = MsoTriState.msoFalse;

                // Create content box
                float contentLeft = 50;
                float contentTop = 100;
                float contentWidth = slide.Design.SlideMaster.Width - 100;
                float contentHeight = slide.Design.SlideMaster.Height - 200;

                PowerPointShape contentShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    contentLeft,
                    contentTop,
                    contentWidth,
                    contentHeight);
                ComReleaser.TrackObject(contentShape);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(contentShape));
                
                contentShape.TextFrame.TextRange.Text = content;
                contentShape.TextFrame.TextRange.Font.Size = 20;
                contentShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51)); // Dark gray
                contentShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                contentShape.TextFrame.WordWrap = MsoTriState.msoTrue;
                contentShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                contentShape.Line.Visible = MsoTriState.msoFalse;

                // Add contact information if provided
                if (!string.IsNullOrEmpty(contactInfo))
                {
                    PowerPointShape contactShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        contentLeft,
                        slide.Design.SlideMaster.Height - 60,
                        contentWidth,
                        30);
                    ComReleaser.TrackObject(contactShape);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(contactShape));
                    
                    contactShape.TextFrame.TextRange.Text = contactInfo;
                    contactShape.TextFrame.TextRange.Font.Size = 14;
                    contactShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                    contactShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
                    contactShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    contactShape.TextFrame.WordWrap = MsoTriState.msoTrue;
                    contactShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                    contactShape.Line.Visible = MsoTriState.msoFalse;
                }

                // Add graphic element (decorative)
                PowerPointShape decorShape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    slide.Design.SlideMaster.Width - 150,
                    contentTop,
                    100,
                    5);
                ComReleaser.TrackObject(decorShape);
                
                decorShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                decorShape.Line.Visible = MsoTriState.msoFalse;

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

                    // Then animate the content
                    Effect contentEffect = slide.TimeLine.MainSequence.AddEffect(
                        contentShape,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(contentEffect);
                    
                    contentEffect.Timing.Duration = 0.7f;

                    // Animate decorative element
                    Effect decorEffect = slide.TimeLine.MainSequence.AddEffect(
                        decorShape,
                        MsoAnimEffect.msoAnimEffectWipe,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    ComReleaser.TrackObject(decorEffect);
                    
                    decorEffect.EffectParameters.Direction = MsoAnimDirection.msoAnimDirectionLeft;
                    decorEffect.Timing.Duration = 0.5f;

                    // Finally animate the contact info if present
                    if (!string.IsNullOrEmpty(contactInfo))
                    {
                        PowerPointShape contactShape = slide.Shapes[slide.Shapes.Count - 1];
                        Effect contactEffect = slide.TimeLine.MainSequence.AddEffect(
                            contactShape,
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        ComReleaser.TrackObject(contactEffect);
                        
                        contactEffect.Timing.Duration = 0.5f;
                    }
                }

                // Add speaker notes if provided
                if (!string.IsNullOrEmpty(notes))
                {
                    slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
                }
                
                // Final cleanup before returning
                ComReleaser.ReleaseOldestObjects(10);

                return slide;
            }
            catch (Exception ex)
            {
                throw new Exception("Error generating summary slide", ex);
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