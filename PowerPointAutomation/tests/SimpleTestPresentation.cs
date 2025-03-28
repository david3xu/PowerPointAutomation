using System;
using System.IO;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAutomation.Utilities;
using System.Runtime.InteropServices;

namespace PowerPointAutomation
{
    /// <summary>
    /// A simple test presentation generator to verify memory optimizations
    /// </summary>
    public class SimpleTestPresentation
    {
        private Application pptApp;
        private Presentation presentation;

        /// <summary>
        /// Generate a simple test presentation to verify memory optimization fixes
        /// </summary>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        public void Generate(string outputPath)
        {
            try
            {
                Console.WriteLine("Creating simple test presentation...");
                
                // Initialize PowerPoint
                pptApp = new Application();
                pptApp.Visible = MsoTriState.msoTrue;
                
                // Track COM objects for later cleanup
                ComReleaser.TrackObject(pptApp);
                
                // Create new presentation
                presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);
                ComReleaser.TrackObject(presentation);
                
                // Set presentation properties
                presentation.PageSetup.SlideSize = PpSlideSizeType.ppSlideSizeOnScreen16x9;
                
                // Add title slide
                Slide titleSlide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutTitle);
                ComReleaser.TrackObject(titleSlide);
                
                // Add title text
                Shape titleShape = titleSlide.Shapes.Title;
                ComReleaser.TrackObject(titleShape);
                titleShape.TextFrame.TextRange.Text = "Memory Optimization Test";
                
                // Add subtitle text
                Shape subtitleShape = titleSlide.Shapes[2]; // Subtitle is usually the second shape
                ComReleaser.TrackObject(subtitleShape);
                subtitleShape.TextFrame.TextRange.Text = "Testing PowerPoint Automation Memory Fixes";
                
                // Add content slide
                Slide contentSlide = presentation.Slides.Add(2, PpSlideLayout.ppLayoutText);
                ComReleaser.TrackObject(contentSlide);
                
                // Add title to content slide
                Shape contentTitle = contentSlide.Shapes.Title;
                ComReleaser.TrackObject(contentTitle);
                contentTitle.TextFrame.TextRange.Text = "Memory Optimization Features";
                
                // Add bullet points
                Shape contentShape = contentSlide.Shapes[2]; // Content placeholder is usually the second shape
                ComReleaser.TrackObject(contentShape);
                contentShape.TextFrame.TextRange.Text = 
                    "• Batch processing of COM objects\n" +
                    "• Age-based COM object tracking and release\n" +
                    "• Incremental presentation generation mode\n" +
                    "• 64-bit process optimization\n" +
                    "• System-level memory optimization scripts\n" +
                    "• Configurable garbage collection settings";
                
                // Add conclusion slide
                Slide conclusionSlide = presentation.Slides.Add(3, PpSlideLayout.ppLayoutTitle);
                ComReleaser.TrackObject(conclusionSlide);
                
                // Add conclusion title
                Shape conclusionTitle = conclusionSlide.Shapes.Title;
                ComReleaser.TrackObject(conclusionTitle);
                conclusionTitle.TextFrame.TextRange.Text = "Memory Optimization Successful!";
                
                // Add conclusion text
                Shape conclusionText = conclusionSlide.Shapes[2];
                ComReleaser.TrackObject(conclusionText);
                conclusionText.TextFrame.TextRange.Text = "The presentation was generated without memory issues.";
                
                // Save the presentation to the specified path
                Console.WriteLine($"Saving presentation to: {outputPath}");
                
                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                presentation.SaveAs(outputPath);
                Console.WriteLine("Presentation saved successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating test presentation: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                CleanupComObjects();
            }
        }
        
        /// <summary>
        /// Clean up COM objects to prevent memory leaks
        /// </summary>
        private void CleanupComObjects()
        {
            Console.WriteLine("Cleaning up COM objects...");
            
            try
            {
                // Close the presentation
                if (presentation != null)
                {
                    try
                    {
                        presentation.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error closing presentation: {ex.Message}");
                    }
                    
                    object presObj = presentation;
                    presentation = null;
                    ComReleaser.ReleaseCOMObject(ref presObj);
                }
                
                // Quit PowerPoint
                if (pptApp != null)
                {
                    try
                    {
                        pptApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error quitting PowerPoint: {ex.Message}");
                    }
                    
                    object appObj = pptApp;
                    pptApp = null;
                    ComReleaser.ReleaseCOMObject(ref appObj);
                }
            }
            finally
            {
                // Release all other tracked COM objects
                ComReleaser.ReleaseAllTrackedObjects();
                
                // Force garbage collection
                ComReleaser.FinalCleanup();
                
                // Check if PowerPoint is still running
                if (ComReleaser.IsProcessRunning("POWERPNT"))
                {
                    Console.WriteLine("Warning: PowerPoint is still running. Attempting to terminate...");
                    int killed = ComReleaser.KillProcess("POWERPNT");
                    Console.WriteLine($"Terminated {killed} PowerPoint process(es).");
                }
            }
        }
    }
} 