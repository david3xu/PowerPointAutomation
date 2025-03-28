using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointAutomation.Utilities;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class for generating core feature slides that showcase key product features
    /// </summary>
    public class CoreFeatureSlide
    {
        // Constants for missing enum values
        private const MsoAnimDirection msoAnimDirectionFromLeft = (MsoAnimDirection)3;
        
        private Presentation presentation;
        private CustomLayout layout;
        
        // Brand colors
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        public CoreFeatureSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a core features slide
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="features">Array of feature texts</param>
        /// <param name="notes">Optional slide notes</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string[] features, string notes = null)
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
                
                // Create background for features
                float contentLeft = 50;
                float contentTop = 100;
                float contentWidth = slide.Design.SlideMaster.Width - 100;
                float contentHeight = slide.Design.SlideMaster.Height - 150;

                PowerPointShape contentBg = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    contentLeft, contentTop, contentWidth, contentHeight);
                ComReleaser.TrackObject(contentBg);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(contentBg));
                
                contentBg.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
                contentBg.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
                contentBg.Line.Weight = 1.0f;
                contentBg.Line.DashStyle = MsoLineDashStyle.msoLineSolid;

                // Send the background to the back so it doesn't overlap other elements
                contentBg.ZOrder(MsoZOrderCmd.msoSendToBack);

                // Add feature items
                int featureCount = features.Length;
                float featureItemHeight = 50;
                float spaceBetween = 20;
                float totalFeaturesHeight = featureCount * featureItemHeight + (featureCount - 1) * spaceBetween;
                float startY = contentTop + (contentHeight - totalFeaturesHeight) / 2;
                
                // Keep track of all feature shapes we create for animations
                List<PowerPointShape> featureBoxes = new List<PowerPointShape>();
                List<PowerPointShape> featureIcons = new List<PowerPointShape>();
                List<PowerPointShape> featureTexts = new List<PowerPointShape>();

                for (int i = 0; i < featureCount; i++)
                {
                    float itemY = startY + i * (featureItemHeight + spaceBetween);
                    
                    // Create feature box
                    PowerPointShape featureBox = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeRoundedRectangle,
                        contentLeft + 30,
                        itemY,
                        contentWidth - 60,
                        featureItemHeight);
                    ComReleaser.TrackObject(featureBox);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(featureBox));
                    
                    featureBoxes.Add(featureBox);

                    // Apply gradient fill
                    featureBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
                    featureBox.Fill.OneColorGradient(
                        MsoGradientStyle.msoGradientHorizontal,
                        1, // Variant
                        0.2f); // Degree
                    featureBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkGray);
                    featureBox.Line.Weight = 1.0f;

                    // Add shadow effect
                    featureBox.Shadow.Visible = MsoTriState.msoTrue;
                    featureBox.Shadow.Type = MsoShadowType.msoShadow6;
                    featureBox.Shadow.Transparency = 0.7f;

                    // Create feature icon (using a simple circle)
                    PowerPointShape icon = slide.Shapes.AddShape(
                        MsoAutoShapeType.msoShapeOval,
                        contentLeft + 45,
                        itemY + (featureItemHeight - 30) / 2,
                        30, 30);
                    ComReleaser.TrackObject(icon);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(icon));
                    
                    featureIcons.Add(icon);
                    
                    icon.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                    icon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                    icon.Line.Weight = 1.0f;

                    // Add feature number
                    icon.TextFrame.TextRange.Text = (i + 1).ToString();
                    icon.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    icon.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                    icon.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    icon.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

                    // Add feature text
                    PowerPointShape featureText = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        contentLeft + 90,
                        itemY + 5,
                        contentWidth - 120,
                        featureItemHeight - 10);
                    ComReleaser.TrackObject(featureText);
                    
                    // Add manual ref to prevent RCW separation
                    Marshal.AddRef(Marshal.GetIUnknownForObject(featureText));
                    
                    featureTexts.Add(featureText);
                    
                    featureText.TextFrame.TextRange.Text = features[i];
                    featureText.TextFrame.TextRange.Font.Size = 18;
                    featureText.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;
                    featureText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51)); // Dark gray
                    featureText.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    featureText.Line.Visible = MsoTriState.msoFalse;
                }

                // Keep all shape lists alive until after animations
                GC.KeepAlive(featureBoxes);
                GC.KeepAlive(featureIcons);
                GC.KeepAlive(featureTexts);

                try
                {
                    // Add animation effects - this part can be skipped if there are issues
                    if (slide.TimeLine != null)
                    {
                        // First animate the background
                        Effect bgEffect = slide.TimeLine.MainSequence.AddEffect(
                            contentBg,
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        ComReleaser.TrackObject(bgEffect);

                        // Then animate each feature with a slight delay
                        for (int i = 0; i < featureCount && i < featureBoxes.Count; i++)
                        {
                            // Get feature box and its elements from our tracked lists
                            PowerPointShape featureBox = featureBoxes[i];
                            PowerPointShape icon = featureIcons[i];
                            PowerPointShape text = featureTexts[i];

                            // Animate feature box
                            Effect boxEffect = slide.TimeLine.MainSequence.AddEffect(
                                featureBox,
                                MsoAnimEffect.msoAnimEffectFly,
                                MsoAnimateByLevel.msoAnimateLevelNone,
                                i == 0 ? MsoAnimTriggerType.msoAnimTriggerAfterPrevious : MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            ComReleaser.TrackObject(boxEffect);
                            
                            boxEffect.EffectParameters.Direction = msoAnimDirectionFromLeft;
                            boxEffect.Timing.Duration = 0.5f;
                            
                            // Set delay between features
                            if (i > 0)
                            {
                                boxEffect.Timing.TriggerDelayTime = 0.2f;
                            }

                            // Animate icon
                            Effect iconEffect = slide.TimeLine.MainSequence.AddEffect(
                                icon,
                                MsoAnimEffect.msoAnimEffectFade,
                                MsoAnimateByLevel.msoAnimateLevelNone,
                                MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            ComReleaser.TrackObject(iconEffect);
                            
                            // Animate text
                            Effect textEffect = slide.TimeLine.MainSequence.AddEffect(
                                text,
                                MsoAnimEffect.msoAnimEffectFade,
                                MsoAnimateByLevel.msoAnimateLevelNone,
                                MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            ComReleaser.TrackObject(textEffect);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not add animations to the core features slide: {ex.Message}");
                    // Continue without animations
                }

                // Add slide notes if provided
                if (!string.IsNullOrEmpty(notes) && slide.NotesPage != null && slide.NotesPage.Shapes.Placeholders.Count > 0)
                {
                    try
                    {
                        PowerPointShape notesShape = slide.NotesPage.Shapes.Placeholders[2];
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
                Console.WriteLine($"Error generating core features slide: {ex.Message}");
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