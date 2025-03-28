using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;

namespace PowerPointAutomation.Utilities
{
    public class PowerPointOperations
    {
        public static void SavePresentation(Presentation presentation, string path)
        {
            SavePresentationInternal(presentation, path);
        }
        
        private static void SavePresentationInternal(Presentation presentation, string path)
        {
            PowerPointApplication app = null;
            
            try
            {
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Console.WriteLine($"Creating directory: {dir}");
                    Directory.CreateDirectory(dir);
                }
                
                app = presentation.Application;
                app.Visible = MsoTriState.msoTrue;
                
                Console.WriteLine($"Saving to file: {Path.GetFullPath(path)}");
                Console.WriteLine($"Directory exists: {Directory.Exists(dir)}");
                
                app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                
                int trackedCount = ComReleaser.GetTrackedObjectCount();
                Console.WriteLine($"Pausing COM object release for save operation (tracked objects: {trackedCount})");
                
                ComReleaser.PauseRelease();
                
                Console.WriteLine("Pausing briefly to stabilize PowerPoint before save...");
                Thread.Sleep(1000);
                
                try
                {
                    int refCount = Marshal.AddRef(Marshal.GetIUnknownForObject(presentation));
                    Console.WriteLine($"Added extra ref count to presentation: {refCount}");
                    
                    Console.WriteLine("Requesting light garbage collection before save...");
                    GC.Collect(0);
                    GC.WaitForPendingFinalizers();
                    
                    Thread.Sleep(500);
                    
                    // Force the save itself
                    Console.WriteLine("Executing SaveAs operation...");
                    string savePath = path;
                    if (!path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                    {
                        savePath = Path.ChangeExtension(path, ".pptx");
                    }
                    
                    // Ensure PowerPoint is ready for save
                    Thread.Sleep(1000);
                    app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
                    
                    bool saveSuccess = false;
                    
                    try
                    {
                        // Primary save attempt
                        presentation.SaveAs(savePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoTrue);
                        Console.WriteLine("SaveAs operation completed");
                        saveSuccess = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Primary save failed: {ex.Message}");
                        
                        // First fallback - try to save via application save command
                        try
                        {
                            Console.WriteLine("Trying first fallback save method...");
                            object missing = Type.Missing;
                            app.ActivePresentation.SaveAs(savePath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoTrue);
                            Console.WriteLine("First fallback save succeeded");
                            saveSuccess = true;
                        }
                        catch (Exception fallbackEx)
                        {
                            Console.WriteLine($"First fallback save failed: {fallbackEx.Message}");
                            
                            // Second fallback - try to use SendKeys to save
                            try
                            {
                                Console.WriteLine("Trying second fallback save method (SendKeys)...");
                                app.Visible = MsoTriState.msoTrue;
                                app.Activate();
                                Thread.Sleep(1000);
                                
                                // Send Alt+F, S (File, Save)
                                System.Windows.Forms.SendKeys.SendWait("%fs");
                                Thread.Sleep(2000);
                                Console.WriteLine("SendKeys save attempted");
                                
                                // Wait to see if file appears
                                for (int i = 0; i < 10; i++)
                                {
                                    if (File.Exists(savePath))
                                    {
                                        saveSuccess = true;
                                        Console.WriteLine("Second fallback save succeeded");
                                        break;
                                    }
                                    Thread.Sleep(1000);
                                }
                            }
                            catch (Exception sendKeysEx)
                            {
                                Console.WriteLine($"Second fallback save failed: {sendKeysEx.Message}");
                            }
                        }
                    }
                    
                    // CRITICAL: Wait for file to be created and fully written
                    Console.WriteLine("Waiting for file to be fully saved...");
                    int maxAttempts = 20; // Increased from 10 to 20
                    int attempt = 0;
                    bool fileReady = false;
                    long lastFileSize = -1;
                    int stableSizeCount = 0;
                    
                    while (!fileReady && attempt < maxAttempts)
                    {
                        try
                        {
                            if (File.Exists(savePath))
                            {
                                using (var fs = File.Open(savePath, FileMode.Open, FileAccess.Read, FileShare.None))
                                {
                                    long currentSize = fs.Length;
                                    
                                    if (currentSize > 0)
                                    {
                                        if (currentSize == lastFileSize)
                                        {
                                            stableSizeCount++;
                                            if (stableSizeCount >= 3) // File size stable for 3 consecutive checks
                                            {
                                                fileReady = true;
                                                Console.WriteLine("File is ready and size is stable");
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            stableSizeCount = 0;
                                            lastFileSize = currentSize;
                                            Console.WriteLine($"File size: {currentSize} bytes");
                                        }
                                    }
                                }
                            }
                        }
                        catch (IOException)
                        {
                            Console.WriteLine($"File not ready yet, attempt {attempt + 1}/{maxAttempts}");
                        }
                        
                        Thread.Sleep(1000); // Wait 1 second between checks
                        attempt++;
                    }
                    
                    if (!fileReady)
                    {
                        Console.WriteLine("WARNING: Could not verify file is fully saved");
                    }
                    
                    // Additional wait to ensure PowerPoint has released the file
                    Thread.Sleep(2000);
                    
                    if (File.Exists(savePath))
                    {
                        var fileInfo = new FileInfo(savePath);
                        Console.WriteLine($"File successfully created at: {savePath}");
                        Console.WriteLine($"File size: {fileInfo.Length} bytes");
                        Console.WriteLine($"Last modified: {fileInfo.LastWriteTime}");
                    }
                    else
                    {
                        Console.WriteLine($"WARNING: File was not created at: {savePath}");
                        
                        try
                        {
                            Console.WriteLine("Attempting alternate save method...");
                            GC.KeepAlive(app);
                            GC.KeepAlive(presentation);
                            
                            object missing = Type.Missing;
                            app.CommandBars.ExecuteMso("FileSaveAs");
                            
                            Thread.Sleep(5000);
                            
                            if (File.Exists(savePath))
                            {
                                Console.WriteLine($"File created through alternate method at: {savePath}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Alternative save approach failed: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    Console.WriteLine($"Resuming COM object tracking");
                    ComReleaser.ResumeRelease();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception during save operation: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                throw;
            }
            finally
            {
                if (app != null)
                {
                    GC.KeepAlive(app);
                }
                GC.KeepAlive(presentation);
                ComReleaser.ResumeRelease();
            }
        }

        public static void TerminatePowerPointProcesses()
        {
            try
            {
                var processes = Process.GetProcessesByName("POWERPNT");
                if (processes.Length > 0)
                {
                    Console.WriteLine($"Found {processes.Length} PowerPoint processes to terminate");
                    
                    foreach (var process in processes)
                    {
                        try
                        {
                            process.CloseMainWindow();
                            if (!process.WaitForExit(2000))
                            {
                                var startInfo = new ProcessStartInfo
                                {
                                    FileName = "taskkill.exe",
                                    Arguments = $"/F /PID {process.Id}",
                                    UseShellExecute = false,
                                    RedirectStandardOutput = true,
                                    RedirectStandardError = true,
                                    CreateNoWindow = true
                                };
                                
                                using (var killProcess = Process.Start(startInfo))
                                {
                                    killProcess.WaitForExit();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not terminate PowerPoint process {process.Id}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex) 
            {
                Console.WriteLine($"Warning: Error checking/terminating PowerPoint processes: {ex.Message}");
            }
        }
    }
} 