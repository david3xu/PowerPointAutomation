using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointAutomation.Slides
{
    /// <summary>
    /// Class responsible for generating title slides with professional formatting
    /// </summary>
    public class TitleSlide
    {
        private Presentation presentation;
        private CustomLayout layout;

        /// <summary>
        /// Initializes a new instance of the TitleSlide class
        /// </summary>
        /// <param name="presentation">The PowerPoint presentation</param>
        /// <param name="layout">The slide layout to use</param>
        public TitleSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a professionally formatted title slide
        /// </summary>
        /// <param name="title">The main title</param>
        /// <param name="subtitle">The subtitle</param>
        /// <param name="presenter">The presenter name</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string subtitle, string presenter = null, string notes = null)
        {
            // Add title slide
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Set title with custom formatting
            Shape titleShape = slide.Shapes.Title;
            titleShape.TextFrame.TextRange.Text = title;
            titleShape.TextFrame.TextRange.Font.Size = 54;
            titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125)); // Dark blue

            // Center the title and add a subtle shadow effect
            titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            titleShape.TextFrame.TextRange.Font.Shadow = MsoTriState.msoTrue;

            // Set subtitle (Shape index 2 is typically the subtitle placeholder in title layouts)
            if (slide.Shapes.Count > 1 && !string.IsNullOrEmpty(subtitle))
            {
                Shape subtitleShape = slide.Shapes[2];
                subtitleShape.TextFrame.TextRange.Text = subtitle;
                subtitleShape.TextFrame.TextRange.Font.Size = 32;
                subtitleShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196)); // Medium blue
                subtitleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            }

            // Add a decorative horizontal line
            float lineWidth = 400;
            float lineTop = slide.Design.SlideMaster.Height * 0.6f;
            Shape lineShape = slide.Shapes.AddLine(
                slide.Design.SlideMaster.Width / 2 - lineWidth / 2, // Start X (centered)
                lineTop, // Start Y
                slide.Design.SlideMaster.Width / 2 + lineWidth / 2, // End X (centered)
                lineTop  // End Y
            );

            // Format line
            lineShape.Line.Weight = 2.0f;
            lineShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(237, 125, 49)); // Orange accent
            lineShape.Line.Style = MsoLineStyle.msoLineSingle;

            // Add presenter name and date if provided
            if (!string.IsNullOrEmpty(presenter))
            {
                // Find a position below the line
                float presenterTop = lineTop + 30;
                Shape presenterShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    slide.Design.SlideMaster.Width / 2 - 200, // Centered
                    presenterTop,
                    400, // Width
                    40  // Height
                );

                presenterShape.TextFrame.TextRange.Text = presenter;
                presenterShape.TextFrame.TextRange.Font.Size = 24;
                presenterShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(89, 89, 89)); // Dark gray
                presenterShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            }

            // Add date
            float dateTop = slide.Design.SlideMaster.Height - 80;
            Shape dateShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                slide.Design.SlideMaster.Width / 2 - 200, // Centered
                dateTop,
                400, // Width
                40  // Height
            );

            dateShape.TextFrame.TextRange.Text = DateTime.Now.ToString("MMMM d, yyyy");
            dateShape.TextFrame.TextRange.Font.Size = 16;
            dateShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.DarkGray);
            dateShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            // Add a subtle logo or graphic
            float logoSize = 80;
            float logoLeft = slide.Design.SlideMaster.Width - logoSize - 40; // Right side with margin
            float logoTop = slide.Design.SlideMaster.Height - logoSize - 40; // Bottom with margin

            // Create a simple "KG" logo using shapes
            Shape logoBackground = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                logoLeft,
                logoTop,
                logoSize,
                logoSize
            );

            // Format logo background
            logoBackground.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125)); // Dark blue
            logoBackground.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196)); // Medium blue
            logoBackground.Line.Weight = 2.0f;

            // Add "KG" text to logo
            Shape logoText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                logoLeft,
                logoTop + logoSize / 2 - 20, // Centered vertically
                logoSize,
                40
            );

            logoText.TextFrame.TextRange.Text = "KG";
            logoText.TextFrame.TextRange.Font.Size = 36;
            logoText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            logoText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            logoText.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            // Add animation effects
            AddEntryAnimations(slide, titleShape, subtitle, presenter, lineShape, dateShape, logoBackground, logoText);

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

            return slide;
        }

        /// <summary>
        /// Adds coordinated entry animations to the elements of the title slide
        /// </summary>
        private void AddEntryAnimations(
            Slide slide,
            Shape titleShape,
            string subtitle,
            string presenter,
            Shape lineShape,
            Shape dateShape,
            Shape logoBackground,
            Shape logoText)
        {
            // Create a sequence of animations

            // 1. Title flies in from top
            Effect titleEffect = slide.TimeLine.MainSequence.AddEffect(
                titleShape,
                MsoAnimEffect.msoAnimEffectFly,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerOnPageClick);

            titleEffect.EffectParameters.Direction = MsoAnimDirection.msoAnimDirectionFromTop;
            titleEffect.Timing.Duration = 1.0f;

            // 2. Subtitle fades in after title
            if (!string.IsNullOrEmpty(subtitle) && slide.Shapes.Count > 1)
            {
                Effect subtitleEffect = slide.TimeLine.MainSequence.AddEffect(
                    slide.Shapes[2],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                subtitleEffect.Timing.Duration = 0.8f;
            }

            // 3. Line wipes from left to right
            Effect lineEffect = slide.TimeLine.MainSequence.AddEffect(
                lineShape,
                MsoAnimEffect.msoAnimEffectWipe,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            lineEffect.EffectParameters.Direction = MsoAnimDirection.msoAnimDirectionFromLeft;
            lineEffect.Timing.Duration = 0.5f;

            // 4. Presenter name fades in after line (if provided)
            if (!string.IsNullOrEmpty(presenter))
            {
                // Find presenter shape (would be shape after the line)
                int presenterShapeIndex = 4; // Typically would be the 4th shape
                if (slide.Shapes.Count >= presenterShapeIndex)
                {
                    Effect presenterEffect = slide.TimeLine.MainSequence.AddEffect(
                        slide.Shapes[presenterShapeIndex],
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    presenterEffect.Timing.Duration = 0.5f;
                }
            }

            // 5. Date fades in with presenter
            Effect dateEffect = slide.TimeLine.MainSequence.AddEffect(
                dateShape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            dateEffect.Timing.Duration = 0.5f;

            // 6. Logo background and text zoom in together
            Effect logoBackgroundEffect = slide.TimeLine.MainSequence.AddEffect(
                logoBackground,
                MsoAnimEffect.msoAnimEffectGrowAndTurn,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

            logoBackgroundEffect.Timing.Duration = 0.7f;

            Effect logoTextEffect = slide.TimeLine.MainSequence.AddEffect(
                logoText,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            logoTextEffect.Timing.Duration = 0.7f;
        }
    }
}