using System;
using System.Collections.Generic;
using System.Drawing;

namespace PowerPointAutomation.Models
{
    /// <summary>
    /// Base class for slide content data models
    /// Provides common properties and methods for all slide types
    /// </summary>
    public abstract class SlideContent
    {
        /// <summary>
        /// The slide title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Speaker notes to be added to the slide
        /// </summary>
        public string Notes { get; set; }

        /// <summary>
        /// Whether to include animations on the slide
        /// </summary>
        public bool IncludeAnimations { get; set; } = true;

        /// <summary>
        /// Default constructor
        /// </summary>
        protected SlideContent() { }

        /// <summary>
        /// Constructor with title
        /// </summary>
        /// <param name="title">The slide title</param>
        protected SlideContent(string title)
        {
            Title = title;
        }

        /// <summary>
        /// Constructor with title and notes
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="notes">Speaker notes</param>
        protected SlideContent(string title, string notes)
        {
            Title = title;
            Notes = notes;
        }
    }

    /// <summary>
    /// Content model for title slide
    /// </summary>
    public class TitleSlideContent : SlideContent
    {
        /// <summary>
        /// Subtitle text
        /// </summary>
        public string Subtitle { get; set; }

        /// <summary>
        /// Presenter name
        /// </summary>
        public string Presenter { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public TitleSlideContent() : base() { }

        /// <summary>
        /// Constructor with title and subtitle
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">The subtitle</param>
        public TitleSlideContent(string title, string subtitle)
            : base(title)
        {
            Subtitle = subtitle;
        }

        /// <summary>
        /// Full constructor
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">The subtitle</param>
        /// <param name="presenter">The presenter name</param>
        /// <param name="notes">Speaker notes</param>
        public TitleSlideContent(string title, string subtitle, string presenter, string notes)
            : base(title, notes)
        {
            Subtitle = subtitle;
            Presenter = presenter;
        }
    }

    /// <summary>
    /// Content model for bullet point slide
    /// </summary>
    public class BulletSlideContent : SlideContent
    {
        /// <summary>
        /// Bullet points to display
        /// </summary>
        public List<string> BulletPoints { get; set; } = new List<string>();

        /// <summary>
        /// Default constructor
        /// </summary>
        public BulletSlideContent() : base() { }

        /// <summary>
        /// Constructor with title
        /// </summary>
        /// <param name="title">The slide title</param>
        public BulletSlideContent(string title) : base(title) { }

        /// <summary>
        /// Constructor with title and bullet points
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="bulletPoints">Array of bullet points</param>
        public BulletSlideContent(string title, string[] bulletPoints)
            : base(title)
        {
            BulletPoints.AddRange(bulletPoints);
        }

        /// <summary>
        /// Full constructor
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="bulletPoints">Array of bullet points</param>
        /// <param name="notes">Speaker notes</param>
        public BulletSlideContent(string title, string[] bulletPoints, string notes)
            : base(title, notes)
        {
            BulletPoints.AddRange(bulletPoints);
        }

        /// <summary>
        /// Adds a bullet point to the slide
        /// </summary>
        /// <param name="bulletPoint">Text of the bullet point</param>
        public void AddBulletPoint(string bulletPoint)
        {
            BulletPoints.Add(bulletPoint);
        }
    }

    /// <summary>
    /// Content model for two-column slide
    /// </summary>
    public class TwoColumnSlideContent : SlideContent
    {
        /// <summary>
        /// Left column bullet points
        /// </summary>
        public List<string> LeftColumnBullets { get; set; } = new List<string>();

        /// <summary>
        /// Right column bullet points
        /// </summary>
        public List<string> RightColumnBullets { get; set; } = new List<string>();

        /// <summary>
        /// Default constructor
        /// </summary>
        public TwoColumnSlideContent() : base() { }

        /// <summary>
        /// Constructor with title
        /// </summary>
        /// <param name="title">The slide title</param>
        public TwoColumnSlideContent(string title) : base(title) { }

        /// <summary>
        /// Constructor with title and bullets for both columns
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="leftColumnBullets">Left column bullet points</param>
        /// <param name="rightColumnBullets">Right column bullet points</param>
        public TwoColumnSlideContent(string title, string[] leftColumnBullets, string[] rightColumnBullets)
            : base(title)
        {
            LeftColumnBullets.AddRange(leftColumnBullets);
            RightColumnBullets.AddRange(rightColumnBullets);
        }

        /// <summary>
        /// Full constructor
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="leftColumnBullets">Left column bullet points</param>
        /// <param name="rightColumnBullets">Right column bullet points</param>
        /// <param name="notes">Speaker notes</param>
        public TwoColumnSlideContent(string title, string[] leftColumnBullets, string[] rightColumnBullets, string notes)
            : base(title, notes)
        {
            LeftColumnBullets.AddRange(leftColumnBullets);
            RightColumnBullets.AddRange(rightColumnBullets);
        }

        /// <summary>
        /// Adds a bullet point to the left column
        /// </summary>
        /// <param name="bulletPoint">Text of the bullet point</param>
        public void AddLeftBulletPoint(string bulletPoint)
        {
            LeftColumnBullets.Add(bulletPoint);
        }

        /// <summary>
        /// Adds a bullet point to the right column
        /// </summary>
        /// <param name="bulletPoint">Text of the bullet point</param>
        public void AddRightBulletPoint(string bulletPoint)
        {
            RightColumnBullets.Add(bulletPoint);
        }
    }

    /// <summary>
    /// Content model for diagram slide
    /// </summary>
    public class DiagramSlideContent : SlideContent
    {
        /// <summary>
        /// Subtitle for the diagram
        /// </summary>
        public string Subtitle { get; set; }

        /// <summary>
        /// Type of diagram to create
        /// </summary>
        public DiagramType DiagramType { get; set; } = DiagramType.KnowledgeGraph;

        /// <summary>
        /// Default constructor
        /// </summary>
        public DiagramSlideContent() : base() { }

        /// <summary>
        /// Constructor with title
        /// </summary>
        /// <param name="title">The slide title</param>
        public DiagramSlideContent(string title) : base(title) { }

        /// <summary>
        /// Constructor with title and subtitle
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">The subtitle</param>
        public DiagramSlideContent(string title, string subtitle)
            : base(title)
        {
            Subtitle = subtitle;
        }

        /// <summary>
        /// Full constructor
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="subtitle">The subtitle</param>
        /// <param name="diagramType">Type of diagram</param>
        /// <param name="notes">Speaker notes</param>
        public DiagramSlideContent(string title, string subtitle, DiagramType diagramType, string notes)
            : base(title, notes)
        {
            Subtitle = subtitle;
            DiagramType = diagramType;
        }
    }

    /// <summary>
    /// Content model for conclusion slide
    /// </summary>
    public class ConclusionSlideContent : SlideContent
    {
        /// <summary>
        /// Main conclusion text
        /// </summary>
        public string ConclusionText { get; set; }

        /// <summary>
        /// "Thank you" text
        /// </summary>
        public string ThankYouText { get; set; } = "Thank You!";

        /// <summary>
        /// Contact information
        /// </summary>
        public string ContactInfo { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ConclusionSlideContent() : base() { }

        /// <summary>
        /// Constructor with title and conclusion text
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="conclusionText">Main conclusion text</param>
        public ConclusionSlideContent(string title, string conclusionText)
            : base(title)
        {
            ConclusionText = conclusionText;
        }

        /// <summary>
        /// Full constructor
        /// </summary>
        /// <param name="title">The slide title</param>
        /// <param name="conclusionText">Main conclusion text</param>
        /// <param name="thankYouText">"Thank you" text</param>
        /// <param name="contactInfo">Contact information</param>
        /// <param name="notes">Speaker notes</param>
        public ConclusionSlideContent(string title, string conclusionText, string thankYouText, string contactInfo, string notes)
            : base(title, notes)
        {
            ConclusionText = conclusionText;
            ThankYouText = thankYouText;
            ContactInfo = contactInfo;
        }
    }

    /// <summary>
    /// Enum for diagram types
    /// </summary>
    public enum DiagramType
    {
        /// <summary>
        /// Knowledge graph diagram showing entities and relationships
        /// </summary>
        KnowledgeGraph,

        /// <summary>
        /// Machine learning integration diagram
        /// </summary>
        MLIntegration,

        /// <summary>
        /// Architectural diagram showing component layers
        /// </summary>
        Architecture
    }
}