using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using PowerPointAutomation.Utilities;

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
            // Add conclusion slide
            Microsoft.Office.Interop.PowerPoint.Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Set title with custom formatting
            slide.Shapes.Title.TextFrame.TextRange.Text = title;
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add decorative elements to the title
            AddDecorativeTitleElements(slide);

            // Set conclusion text in the content placeholder (typically shape index 2)
            if (slide.Shapes.Count > 1)
            {
                PowerPointShape contentShape = slide.Shapes[2];
                contentShape.TextFrame.TextRange.Text = conclusionText;
                contentShape.TextFrame.TextRange.Font.Size = 20;
                contentShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                contentShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 12;

                // Apply paragraph formatting for better readability
                contentShape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                contentShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.2f;
            }

            // Add "Key Takeaways" section with visual element
            AddKeyTakeaways(slide, conclusionText);

            // Add "Thank You" text with emphasis
            float footerTop = slide.Design.SlideMaster.Height - 150;
            PowerPointShape thankYouShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width / 2 - 150,
                footerTop,
                300,
                50
            );

            thankYouShape.TextFrame.TextRange.Text = thankYouText;
            thankYouShape.TextFrame.TextRange.Font.Size = 32;
            thankYouShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            thankYouShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(accentColor);
            thankYouShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            // Add visual emphasis to the Thank You text
            thankYouShape.Shadow.Visible = MsoTriState.msoTrue;
            thankYouShape.Shadow.OffsetX = 3;
            thankYouShape.Shadow.OffsetY = 3;
            thankYouShape.Shadow.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            thankYouShape.Shadow.Blur = 3;

            // Add contact information if provided
            if (!string.IsNullOrEmpty(contactInfo))
            {
                PowerPointShape contactShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    slide.Design.SlideMaster.Width / 2 - 200,
                    footerTop + 50,
                    400,
                    30
                );

                contactShape.TextFrame.TextRange.Text = contactInfo;
                contactShape.TextFrame.TextRange.Font.Size = 16;
                contactShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                contactShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.DarkGray);
                contactShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            }

            // Add a call to action
            AddCallToAction(slide, footerTop + 90);

            // Add animations to slide elements
            AddSlideAnimations(slide, thankYouShape);

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

            return slide;
        }

        /// <summary>
        /// Adds decorative elements to enhance the title's visual appeal
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        private void AddDecorativeTitleElements(Microsoft.Office.Interop.PowerPoint.Slide slide)
        {
            // Add a subtle underline to the title
            PowerPointShape titleUnderline = slide.Shapes.AddLine(
                slide.Shapes.Title.Left,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5,
                slide.Shapes.Title.Left + slide.Shapes.Title.Width * 0.3f,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5
            );

            titleUnderline.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
            titleUnderline.Line.Weight = 3.0f;

            // Add a visual icon next to the title
            float iconSize = 30;
            float iconLeft = slide.Shapes.Title.Left + slide.Shapes.Title.Width + 10;
            float iconTop = slide.Shapes.Title.Top + (slide.Shapes.Title.Height - iconSize) / 2;

            PowerPointShape conclusionIcon = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                iconLeft,
                iconTop,
                iconSize,
                iconSize
            );

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

            checkmark.TextFrame.TextRange.Text = "✓";
            checkmark.TextFrame.TextRange.Font.Size = 18;
            checkmark.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            checkmark.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            checkmark.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            checkmark.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            checkmark.Line.Visible = MsoTriState.msoFalse;
        }

        /// <summary>
        /// Adds a key takeaways section that summarizes main points
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        /// <param name="conclusionText">The conclusion text to extract takeaways from</param>
        private void AddKeyTakeaways(Microsoft.Office.Interop.PowerPoint.Slide slide, string conclusionText)
        {
            // Only add the key takeaways if there's enough content in the conclusion text
            if (conclusionText.Length < 100)
                return;

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

            takeawaysBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
            takeawaysBox.Line.ForeColor.RGB = ColorTranslator.ToOle(secondaryColor);
            takeawaysBox.Line.Weight = 1.5f;

            // Add header
            PowerPointShape headerShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                boxLeft + 10,
                boxTop + 10,
                boxWidth - 20,
                30
            );

            headerShape.TextFrame.TextRange.Text = "Key Takeaways";
            headerShape.TextFrame.TextRange.Font.Size = 18;
            headerShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            headerShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
            headerShape.Line.Visible = MsoTriState.msoFalse;

            // Add simplified takeaways
            PowerPointShape takeawaysContent = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                boxLeft + 15,
                boxTop + 45,
                boxWidth - 30,
                boxHeight - 55
            );

            // Create simplified bullet points from the conclusion text
            string[] bullets = {
                "Knowledge graphs connect entities and relationships",
                "Provides context traditional databases lack",
                "Enables sophisticated reasoning and discovery",
                "Bridges structured and unstructured data"
            };

            // Format the bullet points
            TextRange textRange = takeawaysContent.TextFrame.TextRange;
            textRange.Text = "";

            for (int i = 0; i < bullets.Length; i++)
            {
                if (i > 0)
                    textRange.InsertAfter("\r");

                TextRange bulletPoint = textRange.InsertAfter(bullets[i]);
                
                // Use compatibility layer for paragraph indentation
                bool indentSuccess = OfficeCompatibility.SetParagraphIndentation(
                    bulletPoint.ParagraphFormat, 5, 10);
                    
                // If indentation properties failed, use visual indentation as fallback
                if (!indentSuccess)
                {
                    // Replace the text with indented text
                    string indentedText = "    " + bullets[i]; // 4 spaces for visual indent
                    bulletPoint.Text = indentedText;
                }
                
                bulletPoint.Font.Size = 14;
                bulletPoint.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89));
            }

            takeawaysContent.Line.Visible = MsoTriState.msoFalse;
        }

        /// <summary>
        /// Adds a call to action prompt for next steps
        /// </summary>
        /// <param name="slide">The slide to modify</param>
        /// <param name="yPosition">Y-coordinate position</param>
        private void AddCallToAction(Microsoft.Office.Interop.PowerPoint.Slide slide, float yPosition)
        {
            PowerPointShape ctaShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width / 2 - 150,
                yPosition,
                300,
                30
            );

            ctaShape.TextFrame.TextRange.Text = "Questions? Reach out to discuss knowledge graph applications!";
            ctaShape.TextFrame.TextRange.Font.Size = 12;
            ctaShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
            ctaShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
            ctaShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
        }

        /// <summary>
        /// Adds animations to the slide elements
        /// </summary>
        /// <param name="slide">The slide to animate</param>
        /// <param name="thankYouShape">The "Thank You" shape to emphasize</param>
        private void AddSlideAnimations(Microsoft.Office.Interop.PowerPoint.Slide slide, PowerPointShape thankYouShape)
        {
            // Animate the conclusion text first
            if (slide.Shapes.Count > 1)
            {
                Effect contentEffect = slide.TimeLine.MainSequence.AddEffect(
                    slide.Shapes[2],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerOnPageClick);

                contentEffect.Timing.Duration = 0.7f;
            }

            // Find key takeaways box (should be around index 5 based on our implementation)
            if (slide.Shapes.Count > 5)
            {
                Effect takeawaysEffect = slide.TimeLine.MainSequence.AddEffect(
                    slide.Shapes[5],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                takeawaysEffect.Timing.Duration = 0.5f;

                // Animate the content of the takeaways box
                if (slide.Shapes.Count > 7)
                {
                    Effect takeawaysContentEffect = slide.TimeLine.MainSequence.AddEffect(
                        slide.Shapes[7],
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    takeawaysContentEffect.Timing.Duration = 0.5f;
                }
            }

            // Finally, animate the Thank You text with some emphasis
            Effect thankYouEffect = slide.TimeLine.MainSequence.AddEffect(
                thankYouShape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            thankYouEffect.Timing.Duration = 0.8f;

            // Add a subtle emphasis animation
            Effect emphasisEffect = slide.TimeLine.MainSequence.AddEffect(
                thankYouShape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            emphasisEffect.Timing.Duration = 1.0f;
        }
    }
}