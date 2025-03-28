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
    /// Class for generating introduction slides
    /// </summary>
    public class IntroductionSlide
    {
        // Constants for missing enum values
        private const PpPlaceholderType ppPlaceholderContent = (PpPlaceholderType)2;
        private const MsoAnimDirection msoAnimDirectionFromRight = (MsoAnimDirection)4;
        
        private Presentation presentation;
        private CustomLayout layout;
        
        // Keep track of created COM objects to prevent premature garbage collection
        private List<object> localComReferences = new List<object>();
        
        // Brand colors
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        public IntroductionSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates an introduction slide with title and text content
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="content">The main content text</param>
        /// <param name="notes">Optional slide notes</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string content, string notes = null)
        {
            Slide slide = null;
            try
            {
                Console.WriteLine("Creating introduction slide with enhanced COM object handling...");
                
                // Pause automatic COM object release during slide creation
                ComReleaser.PauseRelease();
                
                // Add slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);
                // Double tracking for this critical object
                ComReleaser.TrackObject(slide);
                localComReferences.Add(slide);
                
                // Add manual ref count to prevent RCW separation
                try {
                    Marshal.AddRef(Marshal.GetIUnknownForObject(slide));
                } catch (Exception ex) {
                    Console.WriteLine($"Warning: Could not add ref to slide: {ex.Message}");
                }

                // Try to get the title shape from the placeholder
                PowerPointShape titleShape = null;
                try
                {
                    titleShape = GetOrCreateTitleShape(slide);
                    ComReleaser.TrackObject(titleShape);
                    localComReferences.Add(titleShape);
                    
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
                    localComReferences.Add(titleShape);
                    
                    // Format the title
                    titleShape.TextFrame.TextRange.Text = title;
                    titleShape.TextFrame.TextRange.Font.Size = 36;
                    titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
                    titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                    titleShape.Line.Visible = MsoTriState.msoFalse;
                }
                
                // DISABLE intermediate cleanup to prevent COM object separation
                // ComReleaser.ReleaseOldestObjects(5);

                // Try to get the content placeholder for body text
                PowerPointShape contentShape = null;
                try
                {
                    // Keep a local copy of the shapes collection to prevent RCW separation
                    var shapesCollection = slide.Shapes;
                    localComReferences.Add(shapesCollection);
                    
                    foreach (PowerPointShape shape in shapesCollection)
                    {
                        // Track each shape to prevent GC
                        ComReleaser.TrackObject(shape);
                        localComReferences.Add(shape);
                        
                        if (shape.Type == MsoShapeType.msoPlaceholder)
                        {
                            if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                            {
                                contentShape = shape;
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not find content placeholder: {ex.Message}");
                }

                // If no content placeholder found, create a custom shape
                if (contentShape == null)
                {
                    // Keep a local reference to the Shapes collection
                    var shapesCollection = slide.Shapes;
                    localComReferences.Add(shapesCollection);
                    
                    contentShape = shapesCollection.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        50, // Left
                        100, // Top
                        slide.Design.SlideMaster.Width - 100, // Width
                        slide.Design.SlideMaster.Height - 150 // Height
                    );
                    ComReleaser.TrackObject(contentShape);
                    localComReferences.Add(contentShape);
                    
                    contentShape.TextFrame.TextRange.Font.Size = 18;
                    contentShape.Line.Visible = MsoTriState.msoFalse;
                }
                else
                {
                    ComReleaser.TrackObject(contentShape);
                    localComReferences.Add(contentShape);
                }

                // Set the content text
                contentShape.TextFrame.TextRange.Text = content;
                contentShape.TextFrame.TextRange.Font.Size = 18;
                contentShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(51, 51, 51)); // Dark gray
                contentShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                contentShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                contentShape.TextFrame.WordWrap = MsoTriState.msoTrue;
                contentShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                contentShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 10;
                contentShape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 5;

                // Add background accent
                PowerPointShape accentShape = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    0, slide.Design.SlideMaster.Height - 40,
                    slide.Design.SlideMaster.Width, 40);
                ComReleaser.TrackObject(accentShape);
                localComReferences.Add(accentShape);
                
                accentShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                accentShape.Line.Visible = MsoTriState.msoFalse;
                accentShape.ZOrder(MsoZOrderCmd.msoSendToBack);

                // Add animation effects
                if (slide.TimeLine != null)
                {
                    var timeLine = slide.TimeLine;
                    localComReferences.Add(timeLine);
                    
                    var mainSequence = timeLine.MainSequence;
                    localComReferences.Add(mainSequence);
                    
                    // First animate the title
                    Effect titleEffect = mainSequence.AddEffect(
                        titleShape,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    ComReleaser.TrackObject(titleEffect);
                    localComReferences.Add(titleEffect);

                    // Then animate the content
                    Effect contentEffect = mainSequence.AddEffect(
                        contentShape,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(contentEffect);
                    localComReferences.Add(contentEffect);
                    
                    contentEffect.Timing.Duration = 0.7f;

                    // Finally animate the accent bar
                    Effect accentEffect = mainSequence.AddEffect(
                        accentShape,
                        MsoAnimEffect.msoAnimEffectWipe,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    ComReleaser.TrackObject(accentEffect);
                    localComReferences.Add(accentEffect);
                    
                    accentEffect.EffectParameters.Direction = MsoAnimDirection.msoAnimDirectionRight;
                    accentEffect.Timing.Duration = 0.5f;
                }

                // Add speaker notes if provided
                if (!string.IsNullOrEmpty(notes))
                {
                    var notesPage = slide.NotesPage;
                    localComReferences.Add(notesPage);
                    
                    var notesShape = notesPage.Shapes[2];
                    localComReferences.Add(notesShape);
                    
                    notesShape.TextFrame.TextRange.Text = notes;
                }
                
                // DISABLE final cleanup to prevent COM object separation
                // ComReleaser.ReleaseOldestObjects(10);
                
                // Keep all local COM references alive until after we return the slide
                foreach (var comObj in localComReferences)
                {
                    GC.KeepAlive(comObj);
                }
                
                // Resume automatic COM object release
                ComReleaser.ResumeRelease();
                
                Console.WriteLine("Introduction slide created successfully with enhanced COM object handling.");

                return slide;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating introduction slide: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                
                // Ensure we resume COM release even on error
                ComReleaser.ResumeRelease();
                
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
                // Keep a local reference to the shapes collection
                var shapesCollection = slide.Shapes;
                localComReferences.Add(shapesCollection);
                
                List<PowerPointShape> shapeList = new List<PowerPointShape>();
                
                foreach (PowerPointShape shape in shapesCollection)
                {
                    // Track each shape to prevent GC
                    ComReleaser.TrackObject(shape);
                    localComReferences.Add(shape);
                    shapeList.Add(shape);
                    
                    if (shape.Type == MsoShapeType.msoPlaceholder)
                    {
                        if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle)
                        {
                            titleShape = shape;
                            break;
                        }
                    }
                }
                
                // Keep shape list alive 
                GC.KeepAlive(shapeList);
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
                
                ComReleaser.TrackObject(titleShape);
                localComReferences.Add(titleShape);
                
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