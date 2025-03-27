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
    /// Class responsible for generating content slides with bullet points
    /// </summary>
    public class ContentSlide
    {
        private Presentation presentation;
        private CustomLayout layout;

        // Theme colors for consistent branding
        private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
        private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
        private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange

        /// <summary>
        /// Initializes a new instance of the ContentSlide class
        /// </summary>
        /// <param name="presentation">The PowerPoint presentation</param>
        /// <param name="layout">The slide layout to use</param>
        public ContentSlide(Presentation presentation, CustomLayout layout)
        {
            this.presentation = presentation;
            this.layout = layout;
        }

        /// <summary>
        /// Generates a content slide with bullet points
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="bulletPoints">Array of bullet point text</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <param name="animateBullets">Whether to animate bullet points</param>
        /// <returns>The created slide</returns>
        public Slide Generate(string title, string[] bulletPoints, string notes = null, bool animateBullets = false)
        {
            // Add content slide
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Set title with custom formatting
            slide.Shapes.Title.TextFrame.TextRange.Text = title;
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add a subtle underline to the title
            PowerPointShape titleUnderline = slide.Shapes.AddLine(
                slide.Shapes.Title.Left,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5,
                slide.Shapes.Title.Left + slide.Shapes.Title.Width * 0.4f,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5
            );

            titleUnderline.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
            titleUnderline.Line.Weight = 3.0f;

            // Access the content placeholder (typically index 2) and add bullet points
            PowerPointShape contentShape = slide.Shapes[2];
            TextRange textRange = contentShape.TextFrame.TextRange;
            textRange.Text = "";

            // Track indentation level (0 = main bullet, 1 = sub-bullet)
            int currentIndentLevel = 0;

            // Add each bullet point with formatting based on indentation
            for (int i = 0; i < bulletPoints.Length; i++)
            {
                string bulletText = bulletPoints[i];

                // Determine indentation level based on leading bullet character
                if (bulletText.StartsWith("• "))
                {
                    currentIndentLevel = 1;
                    bulletText = bulletText.Substring(2); // Remove the bullet character
                }
                else
                {
                    currentIndentLevel = 0;
                }

                // Insert a line break if not the first bullet
                if (i > 0)
                    textRange.InsertAfter("\r");

                // Add the bullet point text
                TextRange newBullet = textRange.InsertAfter(bulletText);

                // Format bullet based on level (customize appearance without using IndentLevel)
                if (currentIndentLevel == 0)
                {
                    // Main bullet - use default bullet formatting
                    newBullet.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                    newBullet.Font.Size = 24;
                    newBullet.Font.Bold = MsoTriState.msoTrue;
                    newBullet.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);
                }
                else
                {
                    // Sub-bullet - indent manually using compatible properties
                    // Apply indentation for sub-bullets
                    // Use FirstLineIndent and LeftIndent instead of First and Left
                    newBullet.ParagraphFormat.FirstLineIndent = 10;
                    newBullet.ParagraphFormat.LeftIndent = 10;
                    newBullet.Font.Size = 20;
                    newBullet.Font.Bold = MsoTriState.msoFalse;
                    newBullet.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
                }

                // Add spacing between bullets
                newBullet.ParagraphFormat.SpaceAfter = 6;
            }

            // Add animations if requested
            if (animateBullets)
            {
                // Animate all bullets at once, since level-by-level animation might not be supported
                Effect bulletEffect = slide.TimeLine.MainSequence.AddEffect(
                    contentShape,
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerOnPageClick);

                // Set timing for bullet animations
                bulletEffect.Timing.Duration = 0.5f;
            }

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

            return slide;
        }

        /// <summary>
        /// Generates a content slide with two columns of bullet points
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="leftColumnBullets">Array of bullet points for left column</param>
        /// <param name="rightColumnBullets">Array of bullet points for right column</param>
        /// <param name="notes">Optional speaker notes</param>
        /// <param name="animateBullets">Whether to animate bullet points</param>
        /// <returns>The created slide</returns>
        public Slide GenerateTwoColumn(
            string title,
            string[] leftColumnBullets,
            string[] rightColumnBullets,
            string notes = null,
            bool animateBullets = false)
        {
            // Add slide with two content placeholders
            Slide slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout);

            // Set title with custom formatting
            slide.Shapes.Title.TextFrame.TextRange.Text = title;
            slide.Shapes.Title.TextFrame.TextRange.Font.Size = 36;
            slide.Shapes.Title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            slide.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(primaryColor);

            // Add a subtle underline to the title
            PowerPointShape titleUnderline = slide.Shapes.AddLine(
                slide.Shapes.Title.Left,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5,
                slide.Shapes.Title.Left + slide.Shapes.Title.Width * 0.4f,
                slide.Shapes.Title.Top + slide.Shapes.Title.Height + 5
            );

            titleUnderline.Line.ForeColor.RGB = ColorTranslator.ToOle(accentColor);
            titleUnderline.Line.Weight = 3.0f;

            // Left column (should be shape index 2 in two-column layout)
            if (slide.Shapes.Count > 1)
            {
                PowerPointShape leftColumn = slide.Shapes[2];
                FormatBulletPoints(leftColumn, leftColumnBullets, primaryColor);

                // Add animation if requested
                if (animateBullets)
                {
                    Effect leftColumnEffect = slide.TimeLine.MainSequence.AddEffect(
                        leftColumn,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerOnPageClick);

                    leftColumnEffect.Timing.Duration = 0.5f;
                }
            }

            // Right column (should be shape index 3 in two-column layout)
            if (slide.Shapes.Count > 2)
            {
                PowerPointShape rightColumn = slide.Shapes[3];
                FormatBulletPoints(rightColumn, rightColumnBullets, secondaryColor);

                // Add animation if requested
                if (animateBullets)
                {
                    Effect rightColumnEffect = slide.TimeLine.MainSequence.AddEffect(
                        rightColumn,
                        MsoAnimEffect.msoAnimEffectFade,
                        MsoAnimateByLevel.msoAnimateLevelNone,
                        MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    rightColumnEffect.Timing.Duration = 0.5f;
                }
            }

            // Add column headings or divider if needed
            float dividerHeight = slide.Design.SlideMaster.Height - 150;
            float dividerY = 120;
            float dividerX = slide.Design.SlideMaster.Width / 2;

            PowerPointShape divider = slide.Shapes.AddLine(
                dividerX, dividerY,
                dividerX, dividerY + dividerHeight - 100
            );

            divider.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGray);
            divider.Line.Weight = 1.5f;
            divider.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;

            // Add speaker notes if provided
            if (!string.IsNullOrEmpty(notes))
            {
                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = notes;
            }

            return slide;
        }

        /// <summary>
        /// Formats the bullet points in a text shape
        /// </summary>
        /// <param name="textShape">The text shape to format</param>
        /// <param name="bulletPoints">Array of bullet point text</param>
        /// <param name="mainColor">Color for main bullet points</param>
        private void FormatBulletPoints(PowerPointShape textShape, string[] bulletPoints, Color mainColor)
        {
            TextRange textRange = textShape.TextFrame.TextRange;
            textRange.Text = "";

            // Track indentation level
            int currentIndentLevel = 0;

            // Add each bullet point with formatting based on indentation
            for (int i = 0; i < bulletPoints.Length; i++)
            {
                string bulletText = bulletPoints[i];

                // Determine indentation level based on leading bullet character
                if (bulletText.StartsWith("• "))
                {
                    currentIndentLevel = 1;
                    bulletText = bulletText.Substring(2); // Remove the bullet character
                }
                else
                {
                    currentIndentLevel = 0;
                }

                // Insert a line break if not the first bullet
                if (i > 0)
                    textRange.InsertAfter("\r");

                // Add the bullet point text
                TextRange newBullet = textRange.InsertAfter(bulletText);

                // Format bullet based on level
                if (currentIndentLevel == 0)
                {
                    // Main bullet - use default bullet formatting
                    newBullet.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                    newBullet.Font.Size = 24;
                    newBullet.Font.Bold = MsoTriState.msoTrue;
                    newBullet.Font.Color.RGB = ColorTranslator.ToOle(mainColor);
                }
                else
                {
                    // Sub-bullet - use compatibility layer for indentation
                    bool indentSuccess = OfficeCompatibility.SetParagraphIndentation(
                        newBullet.ParagraphFormat, 10, 20);
                        
                    // If indentation properties failed, use visual indentation as fallback
                    if (!indentSuccess)
                    {
                        // Replace the text with indented text
                        string indentedText = "    " + bulletText; // 4 spaces for visual indent
                        newBullet.Text = indentedText;
                    }
                    
                    newBullet.Font.Size = 20;
                    newBullet.Font.Bold = MsoTriState.msoFalse;
                    newBullet.Font.Color.RGB = ColorTranslator.ToOle(secondaryColor);
                }

                // Add spacing between bullets
                try
                {
                    newBullet.ParagraphFormat.SpaceAfter = 6;
                }
                catch
                {
                    // Space after not supported in this version - ignore
                }
            }
        }
    }
}
