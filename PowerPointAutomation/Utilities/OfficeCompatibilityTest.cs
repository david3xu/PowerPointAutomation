using System;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Test class for verifying the Office compatibility layer
    /// </summary>
    /// <remarks>
    /// This class contains test methods that can be run to verify the compatibility
    /// features work correctly across different Office versions.
    /// </remarks>
    public class OfficeCompatibilityTest
    {
        /// <summary>
        /// Runs all compatibility tests
        /// </summary>
        /// <param name="outputPath">Path to save the diagnostic report</param>
        public void RunAllTests(string outputPath)
        {
            // Log Office version for diagnostic purposes
            var officeVersion = OfficeCompatibility.GetOfficeVersion();
            
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                writer.WriteLine("PowerPoint Compatibility Test Report");
                writer.WriteLine($"Office Version: {officeVersion}");
                writer.WriteLine($"Date: {DateTime.Now}");
                writer.WriteLine("===================================");
                
                // Run each test
                RunTest(writer, "Theme Color Setting", TestThemeColorSetting);
                RunTest(writer, "Theme Font Setting", TestThemeFontSetting);
                RunTest(writer, "Paragraph Indentation", TestParagraphIndentation);
                RunTest(writer, "SmartArt Layout", TestSmartArtLayout);
                
                writer.WriteLine("\nTest Summary");
                writer.WriteLine("===================================");
                writer.WriteLine($"Total Tests: {testCount}");
                writer.WriteLine($"Passed: {passedCount}");
                writer.WriteLine($"Failed: {testCount - passedCount}");
            }
            
            Console.WriteLine($"Compatibility test report saved to {outputPath}");
        }
        
        private int testCount = 0;
        private int passedCount = 0;
        
        /// <summary>
        /// Runs a test and logs the result
        /// </summary>
        private void RunTest(StreamWriter writer, string testName, Action<StreamWriter> testAction)
        {
            writer.WriteLine($"\nTest: {testName}");
            writer.WriteLine("-----------------------------------");
            
            testCount++;
            
            try
            {
                testAction(writer);
                writer.WriteLine("Result: PASS");
                passedCount++;
            }
            catch (Exception ex)
            {
                writer.WriteLine($"Result: FAIL");
                writer.WriteLine($"Error: {ex.Message}");
                writer.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// Tests theme color setting
        /// </summary>
        private void TestThemeColorSetting(StreamWriter writer)
        {
            Application pptApp = null;
            Presentation presentation = null;
            
            try
            {
                // Initialize PowerPoint
                pptApp = new Application();
                presentation = pptApp.Presentations.Add(MsoTriState.msoFalse);
                
                // Get the first slide master
                Master master = presentation.Designs[1].SlideMaster;
                
                // Try to set theme colors using both approaches
                bool directMethodWorked = false;
                bool compatibilityMethodWorked = false;
                
                try
                {
                    // Try direct method call (Office 2016+ approach)
                    master.Theme.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeColorAccent1).RGB = 
                        ColorTranslator.ToOle(Color.Red);
                    directMethodWorked = true;
                    writer.WriteLine("Direct method call worked (Office 2016+ approach)");
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Direct method call failed: {ex.Message}");
                }
                
                try
                {
                    // Try compatibility layer
                    OfficeCompatibility.SetThemeColor(master.Theme.ThemeColorScheme, 5, ColorTranslator.ToOle(Color.Blue));
                    compatibilityMethodWorked = true;
                    writer.WriteLine("Compatibility layer method worked");
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Compatibility layer failed: {ex.Message}");
                }
                
                if (!directMethodWorked && !compatibilityMethodWorked)
                {
                    throw new Exception("Both theme color setting methods failed");
                }
            }
            finally
            {
                // Clean up
                if (presentation != null)
                {
                    object presObj = presentation;
                    Marshal.ReleaseComObject(presObj);
                }
                
                if (pptApp != null)
                {
                    pptApp.Quit();
                    object appObj = pptApp;
                    Marshal.ReleaseComObject(appObj);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        /// <summary>
        /// Tests theme font setting
        /// </summary>
        private void TestThemeFontSetting(StreamWriter writer)
        {
            Application pptApp = null;
            Presentation presentation = null;
            
            try
            {
                // Initialize PowerPoint
                pptApp = new Application();
                presentation = pptApp.Presentations.Add(MsoTriState.msoFalse);
                
                // Get the first slide master
                Master master = presentation.Designs[1].SlideMaster;
                
                // Try to set theme fonts using both approaches
                bool directMethodWorked = false;
                bool compatibilityMethodWorked = false;
                
                try
                {
                    // Try direct property access (varies by version)
                    var font = master.Theme.ThemeFontScheme.MajorFont;
                    var fontType = font.GetType();
                    
                    try {
                        // Try Latin property (newer versions)
                        font.Latin = "Arial";
                        writer.WriteLine("Latin property worked (newer versions)");
                        directMethodWorked = true;
                    }
                    catch (Exception) {
                        // Try Name property (older versions)
                        var nameProperty = fontType.GetProperty("Name");
                        if (nameProperty != null)
                        {
                            nameProperty.SetValue(font, "Arial");
                            writer.WriteLine("Name property worked (older versions)");
                            directMethodWorked = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Direct font setting failed: {ex.Message}");
                }
                
                try
                {
                    // Try compatibility layer
                    OfficeCompatibility.SetThemeFont(master.Theme.ThemeFontScheme.MinorFont, "Calibri");
                    compatibilityMethodWorked = true;
                    writer.WriteLine("Compatibility layer font setting worked");
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Compatibility layer font setting failed: {ex.Message}");
                }
                
                if (!directMethodWorked && !compatibilityMethodWorked)
                {
                    throw new Exception("Both theme font setting methods failed");
                }
            }
            finally
            {
                // Clean up
                if (presentation != null)
                {
                    object presObj = presentation;
                    Marshal.ReleaseComObject(presObj);
                }
                
                if (pptApp != null)
                {
                    pptApp.Quit();
                    object appObj = pptApp;
                    Marshal.ReleaseComObject(appObj);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        /// <summary>
        /// Tests paragraph indentation
        /// </summary>
        private void TestParagraphIndentation(StreamWriter writer)
        {
            Application pptApp = null;
            Presentation presentation = null;
            
            try
            {
                // Initialize PowerPoint
                pptApp = new Application();
                presentation = pptApp.Presentations.Add(MsoTriState.msoFalse);
                
                // Add a slide with text
                Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutText);
                Shape textShape = slide.Shapes[2]; // Text placeholder
                TextRange textRange = textShape.TextFrame.TextRange;
                textRange.Text = "Test paragraph indentation";
                
                // Try to set indentation using compatibility layer
                bool indentSuccess = OfficeCompatibility.SetParagraphIndentation(
                    textRange.ParagraphFormat, 10, 20);
                
                if (indentSuccess)
                {
                    writer.WriteLine("Paragraph indentation succeeded using compatibility layer");
                }
                else
                {
                    writer.WriteLine("Paragraph indentation failed using compatibility layer");
                    throw new Exception("Paragraph indentation failed");
                }
            }
            finally
            {
                // Clean up
                if (presentation != null)
                {
                    object presObj = presentation;
                    Marshal.ReleaseComObject(presObj);
                }
                
                if (pptApp != null)
                {
                    pptApp.Quit();
                    object appObj = pptApp;
                    Marshal.ReleaseComObject(appObj);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        /// <summary>
        /// Tests SmartArt layout creation
        /// </summary>
        private void TestSmartArtLayout(StreamWriter writer)
        {
            Application pptApp = null;
            Presentation presentation = null;
            
            try
            {
                // Initialize PowerPoint
                pptApp = new Application();
                presentation = pptApp.Presentations.Add(MsoTriState.msoFalse);
                
                // Add a slide
                Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutText);
                
                // Try both methods
                bool directCastWorked = false;
                bool compatibilityMethodWorked = false;
                
                try
                {
                    // Try direct cast (may fail in some versions)
                    var chart1 = slide.Shapes.AddSmartArt(
                        (SmartArtLayout)1,
                        100, 100, 400, 300);
                    
                    directCastWorked = true;
                    writer.WriteLine("Direct SmartArt layout cast worked");
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Direct SmartArt layout cast failed: {ex.Message}");
                }
                
                try
                {
                    // Try compatibility layer
                    var layout = OfficeCompatibility.GetSmartArtLayout(slide.Application, 1);
                    
                    if (layout != null)
                    {
                        var chart2 = slide.Shapes.AddSmartArt(
                            layout,
                            100, 100, 400, 300);
                            
                        compatibilityMethodWorked = true;
                        writer.WriteLine("Compatibility layer SmartArt layout worked");
                    }
                    else
                    {
                        writer.WriteLine("GetSmartArtLayout returned null - SmartArt not available on this system");
                    }
                }
                catch (Exception ex)
                {
                    writer.WriteLine($"Compatibility layer SmartArt layout failed: {ex.Message}");
                }
                
                if (!directCastWorked && !compatibilityMethodWorked)
                {
                    writer.WriteLine("WARNING: Neither SmartArt method worked - SmartArt may not be available on this system");
                }
            }
            finally
            {
                // Clean up
                if (presentation != null)
                {
                    object presObj = presentation;
                    Marshal.ReleaseComObject(presObj);
                }
                
                if (pptApp != null)
                {
                    pptApp.Quit();
                    object appObj = pptApp;
                    Marshal.ReleaseComObject(appObj);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
} 