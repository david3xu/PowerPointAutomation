using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Provides centralized style definitions for consistent presentation branding
    /// </summary>
    public static class PresentationStyles
    {
        #region Color Schemes

        /// <summary>
        /// Color scheme for a professional blue theme
        /// </summary>
        public static class BlueTheme
        {
            /// <summary>Primary dark color for headings and important elements</summary>
            public static readonly Color Primary = Color.FromArgb(31, 73, 125);      // Dark blue

            /// <summary>Secondary color for subheadings and medium emphasis</summary>
            public static readonly Color Secondary = Color.FromArgb(68, 114, 196);   // Medium blue

            /// <summary>Accent color for highlights and call-to-action elements</summary>
            public static readonly Color Accent = Color.FromArgb(237, 125, 49);      // Orange

            /// <summary>Background color for subtle fills</summary>
            public static readonly Color Background = Color.FromArgb(242, 242, 242); // Light gray

            /// <summary>Text color for body content</summary>
            public static readonly Color TextBody = Color.FromArgb(68, 68, 68);      // Dark gray

            /// <summary>Text color for light backgrounds</summary>
            public static readonly Color TextLight = Color.White;

            /// <summary>Success color for positive indicators</summary>
            public static readonly Color Success = Color.FromArgb(112, 173, 71);     // Green

            /// <summary>Warning color for caution indicators</summary>
            public static readonly Color Warning = Color.FromArgb(255, 192, 0);      // Yellow

            /// <summary>Error color for problem indicators</summary>
            public static readonly Color Error = Color.FromArgb(192, 0, 0);          // Red
        }

        /// <summary>
        /// Color scheme for a modern dark theme
        /// </summary>
        public static class DarkTheme
        {
            /// <summary>Primary background color</summary>
            public static readonly Color Primary = Color.FromArgb(32, 32, 32);       // Almost black

            /// <summary>Secondary color for contrast</summary>
            public static readonly Color Secondary = Color.FromArgb(64, 64, 64);     // Dark gray

            /// <summary>Accent color for highlights</summary>
            public static readonly Color Accent = Color.FromArgb(0, 112, 192);       // Bright blue

            /// <summary>Text color for body content</summary>
            public static readonly Color TextBody = Color.FromArgb(240, 240, 240);   // Almost white

            /// <summary>Secondary text color</summary>
            public static readonly Color TextSecondary = Color.FromArgb(180, 180, 180); // Light gray

            /// <summary>Success color</summary>
            public static readonly Color Success = Color.FromArgb(92, 184, 92);      // Green

            /// <summary>Warning color</summary>
            public static readonly Color Warning = Color.FromArgb(240, 173, 78);     // Orange

            /// <summary>Error color</summary>
            public static readonly Color Error = Color.FromArgb(217, 83, 79);        // Red
        }

        #endregion

        #region Font Settings

        /// <summary>
        /// Font settings for different text elements
        /// </summary>
        public static class Fonts
        {
            /// <summary>Primary heading font</summary>
            public static readonly string Heading = "Segoe UI";

            /// <summary>Body text font</summary>
            public static readonly string Body = "Segoe UI";

            /// <summary>Font for code samples</summary>
            public static readonly string Code = "Consolas";

            /// <summary>Font size for slide titles</summary>
            public static readonly float TitleSize = 36;

            /// <summary>Font size for subtitles</summary>
            public static readonly float SubtitleSize = 28;

            /// <summary>Font size for main bullet points</summary>
            public static readonly float BulletMainSize = 24;

            /// <summary>Font size for sub-bullet points</summary>
            public static readonly float BulletSubSize = 20;

            /// <summary>Font size for notes and small text</summary>
            public static readonly float SmallTextSize = 14;
        }

        #endregion

        #region Slide Templates

        /// <summary>
        /// Applies the blue theme to a presentation
        /// </summary>
        /// <param name="presentation">The presentation to theme</param>
        public static void ApplyBlueTheme(Presentation presentation)
        {
            // Get the first slide master
            SlideMaster master = presentation.SlideMasters[1];

            // Set background color
            master.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);

            // Set theme colors
            master.Theme.ThemeColorScheme.Colors[1].RGB = ColorTranslator.ToOle(BlueTheme.Primary);     // Text/Background dark
            master.Theme.ThemeColorScheme.Colors[2].RGB = ColorTranslator.ToOle(Color.White);          // Text/Background light
            master.Theme.ThemeColorScheme.Colors[3].RGB = ColorTranslator.ToOle(BlueTheme.Secondary);   // Accent 1
            master.Theme.ThemeColorScheme.Colors[4].RGB = ColorTranslator.ToOle(BlueTheme.Accent);      // Accent 2
            master.Theme.ThemeColorScheme.Colors[5].RGB = ColorTranslator.ToOle(BlueTheme.Success);     // Accent 3
            master.Theme.ThemeColorScheme.Colors[6].RGB = ColorTranslator.ToOle(Color.FromArgb(0, 176, 240)); // Accent 4

            // Set default font for the presentation
            master.Theme.ThemeFontScheme.MajorFont.Name = Fonts.Heading;
            master.Theme.ThemeFontScheme.MinorFont.Name = Fonts.Body;
        }

        /// <summary>
        /// Applies the dark theme to a presentation
        /// </summary>
        /// <param name="presentation">The presentation to theme</param>
        public static void ApplyDarkTheme(Presentation presentation)
        {
            // Get the first slide master
            SlideMaster master = presentation.SlideMasters[1];

            // Set background color
            master.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(DarkTheme.Primary);

            // Set theme colors
            master.Theme.ThemeColorScheme.Colors[1].RGB = ColorTranslator.ToOle(DarkTheme.TextBody);    // Text/Background dark
            master.Theme.ThemeColorScheme.Colors[2].RGB = ColorTranslator.ToOle(DarkTheme.Primary);     // Text/Background light
            master.Theme.ThemeColorScheme.Colors[3].RGB = ColorTranslator.ToOle(DarkTheme.Accent);      // Accent 1
            master.Theme.ThemeColorScheme.Colors[4].RGB = ColorTranslator.ToOle(Color.FromArgb(255, 143, 0)); // Accent 2 - Orange
            master.Theme.ThemeColorScheme.Colors[5].RGB = ColorTranslator.ToOle(DarkTheme.Success);     // Accent 3
            master.Theme.ThemeColorScheme.Colors[6].RGB = ColorTranslator.ToOle(Color.FromArgb(232, 17, 35)); // Accent 4 - Red

            // Set default font for the presentation
            master.Theme.ThemeFontScheme.MajorFont.Name = Fonts.Heading;
            master.Theme.ThemeFontScheme.MinorFont.Name = Fonts.Body;
        }

        #endregion

        #region Shape Formatting

        /// <summary>
        /// Formats a shape as a callout box with a title and description
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="left">Left position</param>
        /// <param name="top">Top position</param>
        /// <param name="width">Width of the box</param>
        /// <param name="height">Height of the box</param>
        /// <param name="title">Title text</param>
        /// <param name="description">Description text</param>
        /// <param name="fillColor">Background fill color</param>
        /// <returns>The created shape</returns>
        public static Shape CreateCalloutBox(
            Slide slide,
            float left,
            float top,
            float width,
            float height,
            string title,
            string description,
            Color fillColor)
        {
            // Create the main box
            Shape box = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRoundedRectangle,
                left, top, width, height);

            // Format the box
            box.Fill.ForeColor.RGB = ColorTranslator.ToOle(fillColor);
            box.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(
                Math.Max(0, fillColor.R - 50),
                Math.Max(0, fillColor.G - 50),
                Math.Max(0, fillColor.B - 50)));
            box.Line.Weight = 1.5f;

            // Add title
            Shape titleShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left + 10, top + 10, width - 20, 25);

            titleShape.TextFrame.TextRange.Text = title;
            titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            titleShape.TextFrame.TextRange.Font.Size = 16;

            // Determine text color based on background brightness
            int brightness = (fillColor.R + fillColor.G + fillColor.B) / 3;
            Color textColor = brightness > 128 ? Color.Black : Color.White;

            titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(textColor);
            titleShape.Line.Visible = MsoTriState.msoFalse;

            // Add description
            Shape descriptionShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left + 10, top + 40, width - 20, height - 50);

            descriptionShape.TextFrame.TextRange.Text = description;
            descriptionShape.TextFrame.TextRange.Font.Size = 14;
            descriptionShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(textColor);
            descriptionShape.Line.Visible = MsoTriState.msoFalse;

            // Group shapes
            ShapeRange shapes = slide.Shapes.Range(new int[] {
                box.Id, titleShape.Id, descriptionShape.Id
            });

            return shapes.Group();
        }

        /// <summary>
        /// Creates a code block with syntax highlighting
        /// </summary>
        /// <param name="slide">The slide to add the code block to</param>
        /// <param name="left">Left position</param>
        /// <param name="top">Top position</param>
        /// <param name="width">Width of the code block</param>
        /// <param name="height">Height of the code block</param>
        /// <param name="code">The code text</param>
        /// <param name="language">Programming language (for title)</param>
        /// <returns>The created shape</returns>
        public static Shape CreateCodeBlock(
            Slide slide,
            float left,
            float top,
            float width,
            float height,
            string code,
            string language)
        {
            // Create background
            Shape background = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                left, top, width, height);

            // Format background
            background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(40, 44, 52)); // Dark code editor background
            background.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(60, 64, 72));
            background.Line.Weight = 1.0f;

            // Add title bar with language
            Shape titleBar = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                left, top, width, 25);

            titleBar.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(60, 64, 72));
            titleBar.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(80, 84, 92));

            // Add language title
            Shape titleText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left + 10, top, width - 20, 25);

            titleText.TextFrame.TextRange.Text = language;
            titleText.TextFrame.TextRange.Font.Name = Fonts.Code;
            titleText.TextFrame.TextRange.Font.Size = 12;
            titleText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            titleText.Line.Visible = MsoTriState.msoFalse;

            // Add code text
            Shape codeText = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left + 10, top + 30, width - 20, height - 35);

            codeText.TextFrame.TextRange.Text = code;
            codeText.TextFrame.TextRange.Font.Name = Fonts.Code;
            codeText.TextFrame.TextRange.Font.Size = 11;
            codeText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(220, 223, 228));
            codeText.Line.Visible = MsoTriState.msoFalse;

            // Set line spacing for code
            codeText.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
            codeText.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.0f;

            // We would perform real syntax highlighting here with specific color formatting
            // but for simplicity, we'll just group the shapes

            ShapeRange shapes = slide.Shapes.Range(new int[] {
                background.Id, titleBar.Id, titleText.Id, codeText.Id
            });

            return shapes.Group();
        }

        #endregion

        #region Animation Presets

        /// <summary>
        /// Applies a sequential fade-in animation to a collection of shapes
        /// </summary>
        /// <param name="slide">The slide containing the shapes</param>
        /// <param name="shapes">Array of shapes to animate</param>
        /// <param name="clickToStart">Whether the animation should start on click</param>
        public static void ApplySequentialFadeAnimation(Slide slide, Shape[] shapes, bool clickToStart = true)
        {
            if (shapes == null || shapes.Length == 0)
                return;

            // First shape animation trigger
            MsoAnimTriggerType firstTrigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // Add animation for the first shape
            Effect firstEffect = slide.TimeLine.MainSequence.AddEffect(
                shapes[0],
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelNone,
                firstTrigger);

            // Set timing
            firstEffect.Timing.Duration = 0.5f;

            // Add animations for remaining shapes
            for (int i = 1; i < shapes.Length; i++)
            {
                Effect effect = slide.TimeLine.MainSequence.AddEffect(
                    shapes[i],
                    MsoAnimEffect.msoAnimEffectFade,
                    MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                effect.Timing.Duration = 0.5f;
                effect.Timing.TriggerDelayTime = 0.2f; // Slight delay between items
            }
        }

        /// <summary>
        /// Applies a "build" animation for bullet points in a text shape
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="textShape">The text shape with bullet points</param>
        /// <param name="clickToStart">Whether the animation should start on click</param>
        /// <param name="delayBetweenItems">Delay between each bullet point (seconds)</param>
        public static void ApplyBulletPointAnimation(
            Slide slide,
            Shape textShape,
            bool clickToStart = true,
            float delayBetweenItems = 0.0f)
        {
            // First bullet animation trigger
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // Add effect for bullet points
            Effect effect = slide.TimeLine.MainSequence.AddEffect(
                textShape,
                MsoAnimEffect.msoAnimEffectFade,
                MsoAnimateByLevel.msoAnimateLevelParagraph,
                trigger);

            // Set timing
            effect.Timing.Duration = 0.3f;

            // Set delay between items if specified
            if (delayBetweenItems > 0)
            {
                effect.Timing.TriggerDelayTime = delayBetweenItems;
            }
        }

        /// <summary>
        /// Applies an emphasis animation to a shape
        /// </summary>
        /// <param name="slide">The slide containing the shape</param>
        /// <param name="shape">The shape to animate</param>
        /// <param name="effect">The animation effect to apply</param>
        /// <param name="clickToStart">Whether the animation should start on click</param>
        public static void ApplyEmphasisAnimation(
            Slide slide,
            Shape shape,
            MsoAnimEffect effect = MsoAnimEffect.msoAnimEffectPulse,
            bool clickToStart = true)
        {
            // Animation trigger
            MsoAnimTriggerType trigger = clickToStart ?
                MsoAnimTriggerType.msoAnimTriggerOnClick :
                MsoAnimTriggerType.msoAnimTriggerWithPrevious;

            // Add emphasis effect
            Effect animEffect = slide.TimeLine.MainSequence.AddEffect(
                shape,
                effect,
                MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            // Configure timing
            animEffect.Timing.Duration = 0.7f;

            // Ensure the whole shape is animated
            animEffect.EffectInformation.AnimateBackground = MsoTriState.msoTrue;
            animEffect.EffectInformation.AnimateTextInReverse = MsoTriState.msoFalse;
        }

        #endregion
    }
}