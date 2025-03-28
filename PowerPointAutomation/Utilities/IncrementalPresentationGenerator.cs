using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointAutomation.Utilities
{
    public class IncrementalPresentationGenerator
    {
        public static void RunIncrementalPresentation(string finalOutputPath)
        {
            Console.WriteLine("Creating Knowledge Graph presentation in incremental steps...");
            
            string tempDir = Path.Combine(Path.GetTempPath(), "PowerPointTemp");
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            
            foreach (string file in Directory.GetFiles(tempDir, "*.pptx"))
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to delete temporary file {file}: {ex.Message}");
                }
            }
            
            try
            {
                string part1Path = Path.Combine(tempDir, "Part1.pptx");
                Console.WriteLine("Creating Part 1: Title and Introduction...");
                
                string part2Path = Path.Combine(tempDir, "Part2.pptx");
                Console.WriteLine("Creating Part 2: Core Concepts...");
                
                string part3Path = Path.Combine(tempDir, "Part3.pptx");
                Console.WriteLine("Creating Part 3: Applications and Conclusion...");
                
                Console.WriteLine("Merging presentations...");
                MergePresentations(new string[] { part1Path, part2Path, part3Path }, finalOutputPath);
                
                Console.WriteLine($"Incremental presentation creation complete. Final presentation saved at: {finalOutputPath}");
                
                Console.WriteLine("Opening the presentation for review...");
                System.Diagnostics.Process.Start(finalOutputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating incremental presentation: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static void MergePresentations(string[] sourcePaths, string outputPath)
        {
            Application pptApp = null;
            Presentation mainPresentation = null;
            
            try
            {
                pptApp = new Application();
                
                if (File.Exists(sourcePaths[0]))
                {
                    mainPresentation = pptApp.Presentations.Open(
                        sourcePaths[0], MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                }
                else
                {
                    throw new FileNotFoundException($"Source presentation not found: {sourcePaths[0]}");
                }
                
                for (int i = 1; i < sourcePaths.Length; i++)
                {
                    if (File.Exists(sourcePaths[i]))
                    {
                        Console.WriteLine($"Merging presentation {i+1}/{sourcePaths.Length}...");
                        
                        Presentation sourcePresentation = null;
                        
                        try
                        {
                            sourcePresentation = pptApp.Presentations.Open(
                                sourcePaths[i], MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                                
                            int currentSlideCount = mainPresentation.Slides.Count;
                            
                            for (int slideIndex = 1; slideIndex <= sourcePresentation.Slides.Count; slideIndex++)
                            {
                                sourcePresentation.Slides[slideIndex].Copy();
                                mainPresentation.Slides.Paste();
                                
                                if (slideIndex % 5 == 0)
                                {
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                }
                            }
                            
                            sourcePresentation.Close();
                            object sourcePresentationObj = sourcePresentation;
                            sourcePresentation = null;
                            ComReleaser.ReleaseCOMObject(sourcePresentationObj);
                            
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error merging presentation {sourcePaths[i]}: {ex.Message}");
                            
                            if (sourcePresentation != null)
                            {
                                object sourcePresentationObj = sourcePresentation;
                                sourcePresentation = null;
                                ComReleaser.ReleaseCOMObject(sourcePresentationObj);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Source presentation not found: {sourcePaths[i]}");
                    }
                }
                
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                mainPresentation.SaveAs(outputPath);
                
            }
            finally
            {
                if (mainPresentation != null)
                {
                    try
                    {
                        mainPresentation.Close();
                    }
                    catch { }
                    
                    object presentationObj = mainPresentation;
                    mainPresentation = null;
                    ComReleaser.ReleaseCOMObject(presentationObj);
                }
                
                if (pptApp != null)
                {
                    try
                    {
                        pptApp.Quit();
                    }
                    catch { }
                    
                    object appObj = pptApp;
                    pptApp = null;
                    ComReleaser.ReleaseCOMObject(appObj);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
} 