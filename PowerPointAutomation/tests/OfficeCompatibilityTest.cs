using System;
using System.IO;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Class to test and verify compatibility with Microsoft Office COM objects
    /// </summary>
    public class OfficeCompatibilityTest
    {
        /// <summary>
        /// Runs a simple compatibility test to verify PowerPoint automation is working
        /// </summary>
        /// <param name="outputPath">Path to save the test presentation</param>
        /// <returns>True if test passed, false if issues were detected</returns>
        public static bool RunCompatibilityTest(string outputPath)
        {
            Console.WriteLine("Running Office Compatibility Test...");
            
            Application pptApp = null;
            Presentation pres = null;
            bool success = false;
            
            try
            {
                // Create PowerPoint application instance
                pptApp = new Application();
                ComReleaser.TrackObject(pptApp);
                
                // Create a new presentation
                pres = pptApp.Presentations.Add(MsoTriState.msoFalse);
                ComReleaser.TrackObject(pres);
                
                // Add a slide
                CustomLayout layout = pres.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
                ComReleaser.TrackObject(layout);
                
                Slide slide = pres.Slides.AddSlide(1, layout);
                ComReleaser.TrackObject(slide);
                
                // Add title
                Microsoft.Office.Interop.PowerPoint.Shape titleShape = null;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder && 
                        shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle)
                    {
                        titleShape = shape;
                        break;
                    }
                }
                
                if (titleShape != null)
                {
                    ComReleaser.TrackObject(titleShape);
                    titleShape.TextFrame.TextRange.Text = "Office Compatibility Test";
                    titleShape.TextFrame.TextRange.Font.Size = 44;
                    titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                    titleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(31, 73, 125));
                }
                
                // Add a subtitle
                Microsoft.Office.Interop.PowerPoint.Shape subtitleShape = null;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder && 
                        shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSubtitle)
                    {
                        subtitleShape = shape;
                        break;
                    }
                }
                
                if (subtitleShape != null)
                {
                    ComReleaser.TrackObject(subtitleShape);
                    subtitleShape.TextFrame.TextRange.Text = "Test completed successfully on " + DateTime.Now.ToString();
                    subtitleShape.TextFrame.TextRange.Font.Size = 24;
                    subtitleShape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
                }
                
                // Save the presentation
                string folder = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
                
                pres.SaveAs(outputPath);
                Console.WriteLine($"Test presentation saved to: {outputPath}");
                
                // Close the presentation
                pres.Close();
                object presObj = pres;
                pres = null;
                ComReleaser.ReleaseCOMObject(ref presObj);
                
                // Quit PowerPoint
                pptApp.Quit();
                object appObj = pptApp;
                pptApp = null;
                ComReleaser.ReleaseCOMObject(ref appObj);
                
                // Force garbage collection to clean up COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Test was successful
                success = true;
                Console.WriteLine("Office Compatibility Test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Office Compatibility Test failed: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                
                success = false;
            }
            finally
            {
                // Clean up in case of exceptions
                if (pres != null)
                {
                    try
                    {
                        pres.Close();
                        object presObj = pres;
                        pres = null;
                        ComReleaser.ReleaseCOMObject(ref presObj);
                    }
                    catch { }
                }
                
                if (pptApp != null)
                {
                    try
                    {
                        pptApp.Quit();
                        object appObj = pptApp;
                        pptApp = null;
                        ComReleaser.ReleaseCOMObject(ref appObj);
                    }
                    catch { }
                }
                
                // Release all tracked COM objects
                ComReleaser.ReleaseAllTrackedObjects();
                
                // Force final garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Check for PowerPoint processes that might be stuck
                try
                {
                    Process[] processes = Process.GetProcessesByName("POWERPNT");
                    if (processes.Length > 0)
                    {
                        Console.WriteLine($"Warning: {processes.Length} PowerPoint process(es) still running after test.");
                        
                        foreach (Process process in processes)
                        {
                            try
                            {
                                process.Kill();
                                Console.WriteLine("Terminated PowerPoint process.");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to terminate PowerPoint process: {ex.Message}");
                            }
                        }
                    }
                }
                catch { }
            }
            
            return success;
        }

        /// <summary>
        /// Allows the Program.cs class to run the compatibility test
        /// </summary>
        public void RunAllTests()
        {
            string testOutputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "OfficeCompatibilityTest.pptx");
                
            RunCompatibilityTest(testOutputPath);
        }
    }
} 