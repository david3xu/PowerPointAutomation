using System;
using System.Drawing;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointAutomation.Utilities
{
    /// <summary>
    /// Provides cross-version compatibility for Office Interop operations
    /// </summary>
    public static class OfficeCompatibility
    {
        /// <summary>
        /// Sets a theme color safely across different Office versions
        /// </summary>
        /// <param name="colorScheme">The theme color scheme</param>
        /// <param name="colorIndex">The color index (1-12)</param>
        /// <param name="rgb">The RGB color value</param>
        public static void SetThemeColor(ThemeColorScheme colorScheme, int colorIndex, int rgb)
        {
            try
            {
                // Approach 1: Try method call syntax with enum (newer versions)
                var methodInfo = colorScheme.GetType().GetMethod("Colors", new Type[] { typeof(MsoThemeColorSchemeIndex) });
                if (methodInfo != null)
                {
                    // Cast to the enum value that may exist in newer versions
                    var enumValue = (MsoThemeColorSchemeIndex)colorIndex;
                    dynamic color = methodInfo.Invoke(colorScheme, new object[] { enumValue });
                    color.RGB = rgb;
                    return;
                }
            }
            catch
            {
                // Method approach failed, silent fallthrough to next approach
            }

            try
            {
                // Approach 2: Try indexer syntax (works with most versions)
                // First, get the Colors property (which might return a collection)
                var colorsProperty = colorScheme.GetType().GetProperty("Colors");
                if (colorsProperty != null)
                {
                    // Get the collection object
                    var colorsCollection = colorsProperty.GetValue(colorScheme);
                    
                    // Try to access the indexer through the collection's Item property
                    var itemProperty = colorsCollection.GetType().GetProperty("Item");
                    if (itemProperty != null)
                    {
                        // Get the color at the specified index
                        var color = itemProperty.GetValue(colorsCollection, new object[] { colorIndex });
                        
                        // Set the RGB value using reflection
                        var rgbProperty = color.GetType().GetProperty("RGB");
                        if (rgbProperty != null)
                        {
                            rgbProperty.SetValue(color, rgb);
                            return;
                        }
                    }
                }
                
                // If we get here, we couldn't use the indexer approach either
                throw new Exception("Could not set theme color using indexer");
            }
            catch (Exception ex)
            {
                // Both approaches failed - log the error
                Console.WriteLine($"Could not set theme color: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets a theme font safely across different Office versions
        /// </summary>
        /// <param name="font">The theme font to modify</param>
        /// <param name="fontName">The font name to set</param>
        public static void SetThemeFont(ThemeFonts font, string fontName)
        {
            try
            {
                // Try Name property first (older versions)
                var nameProperty = font.GetType().GetProperty("Name");
                if (nameProperty != null)
                {
                    nameProperty.SetValue(font, fontName);
                    return;
                }
            }
            catch
            {
                // Name property approach failed, silent fallthrough to next approach
            }

            try
            {
                // Try Latin property (newer versions) using reflection
                var latinProperty = font.GetType().GetProperty("Latin");
                if (latinProperty != null)
                {
                    latinProperty.SetValue(font, fontName);
                    return;
                }
            }
            catch (Exception ex)
            {
                // Both approaches failed - log the error
                Console.WriteLine($"Could not set theme font: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets paragraph indentation safely across different Office versions
        /// </summary>
        /// <param name="format">The paragraph format to modify</param>
        /// <param name="firstLineIndent">First line indent value</param>
        /// <param name="leftIndent">Left indent value</param>
        /// <returns>True if successful, false if fallback needed</returns>
        public static bool SetParagraphIndentation(ParagraphFormat format, float firstLineIndent, float leftIndent)
        {
            try
            {
                // Approach 1: Try newer property names
                var firstProperty = format.GetType().GetProperty("FirstLineIndent");
                var leftProperty = format.GetType().GetProperty("LeftIndent");
                
                bool success = false;
                
                if (firstProperty != null)
                {
                    firstProperty.SetValue(format, firstLineIndent);
                    success = true;
                }
                    
                if (leftProperty != null)
                {
                    leftProperty.SetValue(format, leftIndent);
                    success = true;
                }
                
                if (success)
                    return true;
            }
            catch
            {
                // First approach failed, silent fallthrough to next approach
            }

            try
            {
                // Approach 2: Try older property names
                var firstProperty = format.GetType().GetProperty("First");
                var leftProperty = format.GetType().GetProperty("Left");
                
                bool success = false;
                
                if (firstProperty != null)
                {
                    firstProperty.SetValue(format, firstLineIndent);
                    success = true;
                }
                    
                if (leftProperty != null)
                {
                    leftProperty.SetValue(format, leftIndent);
                    success = true;
                }
                
                return success;
            }
            catch
            {
                // Both approaches failed
                return false;
            }
        }

        /// <summary>
        /// Gets a SmartArt layout safely across different Office versions
        /// </summary>
        /// <param name="application">The PowerPoint application</param>
        /// <param name="index">The layout index (1-based)</param>
        /// <returns>The SmartArt layout object or null if unavailable</returns>
        public static object GetSmartArtLayout(Application application, int index)
        {
            try
            {
                // Try to get layout from the application's collection
                return application.SmartArtLayouts[index];
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"Could not get SmartArt layout: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets the installed Office version
        /// </summary>
        /// <returns>The Office version or null if not detected</returns>
        public static Version GetOfficeVersion()
        {
            try
            {
                // Try to get version from PowerPoint registry key
                using (var key = Registry.ClassesRoot.OpenSubKey("PowerPoint.Application\\CurVer"))
                {
                    if (key != null)
                    {
                        string version = key.GetValue(null) as string;
                        if (!string.IsNullOrEmpty(version))
                        {
                            // Parse version string (e.g., "PowerPoint.Application.16" for Office 2016)
                            string[] parts = version.Split('.');
                            if (parts.Length > 0)
                            {
                                int majorVersion = int.Parse(parts[parts.Length - 1]);
                                
                                // Map version number to office version
                                switch (majorVersion)
                                {
                                    case 14: return new Version(14, 0); // Office 2010
                                    case 15: return new Version(15, 0); // Office 2013
                                    case 16:
                                        // Version 16 could be 2016, 2019, or 365 - need additional checks
                                        return new Version(16, 0);
                                    default:
                                        return new Version(majorVersion, 0);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error detecting Office version: {ex.Message}");
            }
            
            // Default to Office 2013 if detection fails
            return new Version(15, 0);
        }

        /// <summary>
        /// Adds diagnostic logging for Office operations
        /// </summary>
        /// <param name="operationName">Name of the operation</param>
        /// <param name="action">The action to perform</param>
        public static void LogOperation(string operationName, Action action)
        {
            Console.WriteLine($"Starting: {operationName}");
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                action();
                stopwatch.Stop();
                Console.WriteLine($"Completed: {operationName} in {stopwatch.ElapsedMilliseconds}ms");
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Console.WriteLine($"Error in {operationName} after {stopwatch.ElapsedMilliseconds}ms: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw the exception
            }
        }

        /// <summary>
        /// Safely gets the title shape from a slide or creates a custom title if it doesn't exist
        /// </summary>
        /// <param name="slide">The slide to get or create a title for</param>
        /// <param name="title">The title text to set</param>
        /// <param name="fontSize">The font size for the title</param>
        /// <param name="primaryColor">The primary color for the title</param>
        /// <returns>The title shape, either existing or newly created</returns>
        public static PowerPointShape GetOrCreateTitleShape(Slide slide, string title, float fontSize, int primaryColorRgb)
        {
            PowerPointShape titleShape;
            float titleLeft = 50;
            float titleTop = 20;
            float titleWidth = slide.Design.SlideMaster.Width - 100;
            float titleHeight = 50;

            try
            {
                // Try to use the existing title placeholder
                titleShape = slide.Shapes.Title;
                
                // Get position properties for layout consistency if needed later
                titleLeft = titleShape.Left;
                titleTop = titleShape.Top;
                titleWidth = titleShape.Width;
                titleHeight = titleShape.Height;
            }
            catch
            {
                // Title placeholder doesn't exist, create a custom title
                Console.WriteLine("Creating custom title shape for slide");
                titleShape = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    titleLeft, titleTop, titleWidth, titleHeight);
            }

            // Set standard title properties
            titleShape.TextFrame.TextRange.Text = title;
            titleShape.TextFrame.TextRange.Font.Size = fontSize;
            titleShape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            titleShape.TextFrame.TextRange.Font.Color.RGB = primaryColorRgb;
            
            return titleShape;
        }
    }
} 