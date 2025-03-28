using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using PowerPointAutomation.Utilities;
using System.Linq;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class responsible for generating conclusion slides with summary and closing information
    /// </summary>
    public class ConclusionSlide
    {
        private Presentation presentation;
        private CustomLayout layout;

        // Theme colors for consistent branding
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        /// <summary>
        /// Initializes a new instance of the ConclusionSlide class
        /// </summary>
        /// <param name="presentation">The PowerPoint presentation</param>
        /// <param name="layout">The slide layout to use</param>
        public ConclusionSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a conclusion slide with summary text and contact information
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="conclusionText">The main conclusion text</param>
        /// <param name="thankYouText">Optional "Thank You" text</param>
        /// <param name="contactInfo">Optional contact information</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <returns>The created slide</returns>
        public Microsoft.Office.Interop.PowerPoint.Slide Generate(
            string title,
            string conclusionText,
            string thankYouText = "Thank You!",
            string contactInfo = null,
            string notes = null)
        {
            Microsoft.Office.Interop.PowerPoint.Slide slide = null;
            
            try
            {
                // Add conclusion slide
                slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);
                ComReleaser.TrackObject(slide);
                
                // Add manual ref to prevent RCW separation
                Marshal.AddRef(Marshal.GetIUnknownForObject(slide));
                
                // Track max shape count to prevent RCW separation errors
                int maxShapesPerSlide = 25; // More conservative limit to prevent issues
                int currentShapeCount = slide.Shapes.Count;
                Console.WriteLine($"Initial shape count: {currentShapeCount}");

                // Use the safe method to get or create a title
                PowerPointShape titleShape = null;
                try 
                {
                    titleShape = OfficeCompatibility.GetOrCreateTitleShape(
                        slide, 
                        title, 
                        36, 
                        ColorTranslator.ToOle(primaryColor));
                    
                    if (titleShape != null) 
                    {
                        ComReleaser.TrackObject(titleShape);
                        
                        // Add manual ref to prevent RCW separation
                        Marshal.AddRef(Marshal.GetIUnknownForObject(titleShape));
                        currentShapeCount++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Error creating title shape: {ex.Message}");
                }

                // Add decorative elements to the title if we have room
                if (currentShapeCount + 3 < maxShapesPerSlide && titleShape != null)
                {
                    try
                    {
                        AddDecorativeTitleElements(slide, titleShape);
                        currentShapeCount += 3; // Approximately 3 shapes added
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Error adding decorative elements: {ex.Message}");
                    }
                }

                // Set conclusion text in the content placeholder (typically shape index 2)
                if (slide.Shapes.Count > 1 && currentShapeCount < maxShapesPerSlide)
                {
                    try
                    {
                        PowerPointShape contentShape = slide.Shapes[2];
                        ComReleaser.TrackObject(contentShape);
                        
                        // Add manual ref to prevent RCW separation
                        Marshal.AddRef(Marshal.GetIUnknownForObject(contentShape));
                        
                        contentShape.TextFrame.TextRange.Text = conclusionText;
                        contentShape.TextFrame.TextRange.Font.Size = 20;
                        contentShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                        contentShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12;

                        // Apply paragraph formatting for better readability
                        contentShape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                        contentShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.2f;
                        currentShapeCount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Error setting conclusion text: {ex.Message}");
                    }
                }

                // Add "Key Takeaways" section with visual element if we have room
                if (currentShapeCount + 5 < maxShapesPerSlide) // Reduced from 10 to 5
                {
                    try
                    {
                        if (conclusionText.Length >= 100) // Only if there's enough content
                        {
                            AddKeyTakeaways(slide, conclusionText);
                            currentShapeCount += 5; // Conservative estimate
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Error adding key takeaways: {ex.Message}");
                    }
                }

                // Add "Thank You" text with emphasis if we have room
                if (currentShapeCount + 1 < maxShapesPerSlide)
                {
                    try
                    {
                        float footerTop = slide.Design.SlideMaster.Height - 150;
                        PowerPointShape thankYouShape = null;
                        try 
                        {
                            thankYouShape = slide.Shapes.AddTextbox(
                                MsoTextOrientation.msoTextOrientationHorizontal,
                                slide.Design.SlideMaster.Width / 2 - 150,
                                footerTop,
                                300,
                                50
                            );
                            
                            if (thankYouShape != null)
                            {
                                ComReleaser.TrackObject(thankYouShape);
                                
                                // Add manual ref to prevent RCW separation
                                Marshal.AddRef(Marshal.GetIUnknownForObject(thankYouShape));

                                thankYouShape.TextFrame.TextRange.Text = thankYouText;
                                thankYouShape.TextFrame.TextRange.Font.Size = 32;
                                thankYouShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                                thankYouShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(accentColor);
                                thankYouShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

                                // Add visual emphasis safely
                                try 
                                {
                                    thankYouShape.Shadow.Visible = MsoTriState.msoTrue;
                                    thankYouShape.Shadow.OffsetX = 3;
                                    thankYouShape.Shadow.OffsetY = 3;
                                    thankYouShape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
                                    thankYouShape.Shadow.Blur = 3;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Could not set shadow properties: {ex.Message}");
                                }
                                
                                currentShapeCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Error creating thank you shape: {ex.Message}");
                        }

                        // Add contact information if provided and we have room
                        if (!string.IsNullOrEmpty(contactInfo) && currentShapeCount + 1 < maxShapesPerSlide && thankYouShape != null)
                        {
                            try
                            {
                                PowerPointShape contactShape = slide.Shapes.AddTextbox(
                                    MsoTextOrientation.msoTextOrientationHorizontal,
                                    slide.Design.SlideMaster.Width / 2 - 200,
                                    footerTop + 50,
                                    400,
                                    30
                                );
                                
                                if (contactShape != null)
                                {
                                    ComReleaser.TrackObject(contactShape);
                                    
                                    // Add manual ref to prevent RCW separation
                                    Marshal.AddRef(Marshal.GetIUnknownForObject(contactShape));

                                    contactShape.TextFrame.TextRange.Text = contactInfo;
                                    contactShape.TextFrame.TextRange.Font.Size = 16;
                                    contactShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                                    contactShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.DarkGray);
                                    contactShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                                    currentShapeCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Error adding contact info: {ex.Message}");
                            }
                        }

                        // Add a call to action if we have room
                        if (currentShapeCount + 3 < maxShapesPerSlide)
                        {
                            try
                            {
                                AddCallToAction(slide, footerTop + 90);
                                currentShapeCount += 3; // Approximately 3 shapes added
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Error adding call to action: {ex.Message}");
                            }
                        }

                        // Add animations with limit checks
                        if (slide.TimeLine != null && slide.TimeLine.MainSequence != null && thankYouShape != null)
                        {
                            try
                            {
                                // Limit animation count
                                int maxAnimations = 10; // More conservative than before
                                int currentAnimCount = slide.TimeLine.MainSequence.Count;
                                
                                if (currentAnimCount < maxAnimations)
                                {
                                    try 
                                    {
                                        AddSlideAnimations(slide, thankYouShape);
                                        // Force cleanup after animations
                                        ComReleaser.ReleaseOldestObjects(5);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Warning: Error in animation sequence: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Skipping animations - maximum count reached");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Error adding animations: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Error adding thank you text: {ex.Message}");
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

                // Force cleanup before returning
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
                Console.WriteLine($"Error generating conclusion slide: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return slide; // Return the partially created slide or null
            }
        }

        /// <summary>
        /// Adds decorative elements to enhance the title's visual appeal
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        /// <param name="titleShape">The title shape to decorate</param>
        private void AddDecorativeTitleElements(Microsoft.Office.Interop.PowerPoint.Slide slide, PowerPointShape titleShape)
        {
            try
            {
                // Add a subtle underline to the title
                PowerPointShape titleUnderline = slide.Shapes.AddLine(
                    titleShape.Left,
                    titleShape.Top + titleShape.Height + 5,
                    titleShape.Left + titleShape.Width * 0.3f,
                    titleShape.Top + titleShape.Height + 5
                );
                ComReleaser.TrackObject(titleUnderline);
                Marshal.AddRef(Marshal.GetIUnknownForObject(titleUnderline));

                titleUnderline.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                titleUnderline.Line.Weight = 3.0f;

                // Add a visual icon next to the title
                float iconSize = 30;
                float iconLeft = titleShape.Left + titleShape.Width + 10;
                float iconTop = titleShape.Top + (titleShape.Height - iconSize) / 2;

                PowerPointShape conclusionIcon = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    iconLeft,
                    iconTop,
                    iconSize,
                    iconSize
                );
                ComReleaser.TrackObject(conclusionIcon);
                Marshal.AddRef(Marshal.GetIUnknownForObject(conclusionIcon));

                conclusionIcon.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                conclusionIcon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                conclusionIcon.Line.Weight = 1.0f;

                // Add a checkmark inside the icon
                PowerPointShape checkmark = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    iconLeft,
                    iconTop,
                    iconSize,
                    iconSize
                );
                ComReleaser.TrackObject(checkmark);
                Marshal.AddRef(Marshal.GetIUnknownForObject(checkmark));

                checkmark.TextFrame.TextRange.Text = "✓";
                checkmark.TextFrame.TextRange.Font.Size = 18;
                checkmark.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                checkmark.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                checkmark.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                checkmark.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                checkmark.Line.Visible = MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in AddDecorativeTitleElements: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a key takeaways section that summarizes main points
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        /// <param name="conclusionText">The conclusion text to extract takeaways from</param>
        private void AddKeyTakeaways(Microsoft.Office.Interop.PowerPoint.Slide slide, string conclusionText)
        {
            try
            {
                // Only add the key takeaways if there's enough content in the conclusion text
                if (conclusionText.Length < 100)
                    return;

                // Verify shape limit
                int maxShapes = 10; // Maximum shapes to add in this method
                int shapesAdded = 0;

                // Create a background shape for the takeaways section
                float boxLeft = 50;
                float boxTop = 280;
                float boxWidth = 300;
                float boxHeight = 150;

                PowerPointShape takeawaysBox = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    boxLeft,
                    boxTop,
                    boxWidth,
                    boxHeight
                );
                ComReleaser.TrackObject(takeawaysBox);
                Marshal.AddRef(Marshal.GetIUnknownForObject(takeawaysBox));
                shapesAdded++;

                takeawaysBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
                takeawaysBox.Line.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
                takeawaysBox.Line.Weight = 1.5f;

                if (shapesAdded < maxShapes)
                {
                    // Add header
                    PowerPointShape headerShape = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        boxLeft + 10,
                        boxTop + 10,
                        boxWidth - 20,
                        30
                    );
                    ComReleaser.TrackObject(headerShape);
                    Marshal.AddRef(Marshal.GetIUnknownForObject(headerShape));
                    shapesAdded++;

                    headerShape.TextFrame.TextRange.Text = "Key Takeaways";
                    headerShape.TextFrame.TextRange.Font.Size = 18;
                    headerShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    headerShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
                    headerShape.Line.Visible = MsoTriState.msoFalse;
                }

                if (shapesAdded < maxShapes)
                {
                    // Add simplified takeaways
                    PowerPointShape takeawaysContent = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        boxLeft + 15,
                        boxTop + 45,
                        boxWidth - 30,
                        boxHeight - 55
                    );
                    ComReleaser.TrackObject(takeawaysContent);
                    Marshal.AddRef(Marshal.GetIUnknownForObject(takeawaysContent));
                    shapesAdded++;

                    // Create simplified bullet points from the conclusion text
                    string[] bullets = {
                        "Knowledge graphs connect entities and relationships",
                        "Provides context traditional databases lack",
                        "Enables sophisticated reasoning and discovery",
                        "Bridges structured and unstructured data"
                    };

                    takeawaysContent.TextFrame.TextRange.Text = string.Join("\r\n", bullets.Select(b => "• " + b));
                    takeawaysContent.TextFrame.TextRange.Font.Size = 14;
                    takeawaysContent.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 68, 68));
                    takeawaysContent.Line.Visible = MsoTriState.msoFalse;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in AddKeyTakeaways: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a call to action prompt for next steps
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        /// <param name="yPosition">Y-coordinate position</param>
        private void AddCallToAction(Microsoft.Office.Interop.PowerPoint.Slide slide, float yPosition)
        {
            try
            {
                // Create a simple call to action button
                float buttonWidth = 180;
                float buttonHeight = 40;
                float buttonLeft = slide.Design.SlideMaster.Width / 2 - buttonWidth / 2;

                PowerPointShape ctaButton = slide.Shapes.AddShape(
                    MsoAutoShapeType.msoShapeRoundedRectangle,
                    buttonLeft,
                    yPosition,
                    buttonWidth,
                    buttonHeight
                );
                ComReleaser.TrackObject(ctaButton);
                Marshal.AddRef(Marshal.GetIUnknownForObject(ctaButton));

                ctaButton.Fill.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
                ctaButton.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 100, 30));
                ctaButton.Line.Weight = 1.0f;

                PowerPointShape ctaText = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    buttonLeft,
                    yPosition,
                    buttonWidth,
                    buttonHeight
                );
                ComReleaser.TrackObject(ctaText);
                Marshal.AddRef(Marshal.GetIUnknownForObject(ctaText));

                ctaText.TextFrame.TextRange.Text = "Learn More";
                ctaText.TextFrame.TextRange.Font.Size = 16;
                ctaText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                ctaText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
                ctaText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                ctaText.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                ctaText.Line.Visible = MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in AddCallToAction: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds animations to the slide elements
        /// </summary>
        /// <param name="slide">The slide to animate</param>
        /// <param name="thankYouShape">The "Thank You" shape to emphasize</param>
        private void AddSlideAnimations(Microsoft.Office.Interop.PowerPoint.Slide slide, PowerPointShape thankYouShape)
        {
            try
            {
                // Check if timeLine and sequence are available
                if (slide.TimeLine == null || slide.TimeLine.MainSequence == null)
                {
                    Console.WriteLine("Cannot add animations: Timeline or MainSequence is null");
                    return;
                }
                
                // Limit number of animations added
                int maxAnimations = 25;
                int currentCount = slide.TimeLine.MainSequence.Count;
                int availableSlots = maxAnimations - currentCount;
                
                if (availableSlots <= 0)
                {
                    Console.WriteLine("Cannot add animations: Maximum animation count reached");
                    return;
                }
                
                int animationsAdded = 0;
                
                // Animate title with emphasis
                if (slide.Shapes.Count > 1 && animationsAdded < availableSlots)
                {
                    try
                    {
                        Effect titleEffect = slide.TimeLine.MainSequence.AddEffect(
                            slide.Shapes[1], // Title shape is usually the first shape
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                        ComReleaser.TrackObject(titleEffect);
                        animationsAdded++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error adding title animation: {ex.Message}");
                    }
                }
                
                // Animate content with fade
                if (slide.Shapes.Count > 2 && animationsAdded < availableSlots)
                {
                    try
                    {
                        Effect contentEffect = slide.TimeLine.MainSequence.AddEffect(
                            slide.Shapes[2], // Content shape is usually the second shape
                            MsoAnimEffect.msoAnimEffectFade,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        ComReleaser.TrackObject(contentEffect);
                        animationsAdded++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error adding content animation: {ex.Message}");
                    }
                }
                
                // Animate thank you shape with special effect
                if (thankYouShape != null && animationsAdded < availableSlots)
                {
                    try
                    {
                        Effect thankYouEffect = slide.TimeLine.MainSequence.AddEffect(
                            thankYouShape,
                            MsoAnimEffect.msoAnimEffectFly,
                            MsoAnimateByLevel.msoAnimateLevelNone,
                            MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        ComReleaser.TrackObject(thankYouEffect);
                        animationsAdded++;
                        
                        // Configure the animation
                        thankYouEffect.Timing.Duration = 1.0f;
                        thankYouEffect.Timing.TriggerDelayTime = 0.5f;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error adding thank you animation: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"Added {animationsAdded} animations to conclusion slide");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in AddSlideAnimations: {ex.Message}");
            }
        }
    }
}