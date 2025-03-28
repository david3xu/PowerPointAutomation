using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

class PowerPointTest
{
    static void Main()
    {
        Console.WriteLine("PowerPoint COM Test");
        Console.WriteLine("===================");
        
        Application pptApp = null;
        
        try
        {
            Console.WriteLine("Attempting to create PowerPoint application...");
            pptApp = new Application();
            
            Console.WriteLine("PowerPoint application created successfully.");
            Console.WriteLine($"PowerPoint version: {pptApp.Version}");
            
            // Make PowerPoint visible
            pptApp.Visible = MsoTriState.msoTrue;
            Console.WriteLine("Made PowerPoint visible.");
            
            // Create a new presentation
            Console.WriteLine("Creating a new presentation...");
            Presentation pres = pptApp.Presentations.Add(MsoTriState.msoTrue);
            
            Console.WriteLine("New presentation created.");
            Console.WriteLine($"Slide count: {pres.Slides.Count}");
            
            // Close PowerPoint
            Console.WriteLine("Closing PowerPoint...");
            pptApp.Quit();
            
            Console.WriteLine("Test completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
            }
        }
        finally
        {
            if (pptApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(pptApp);
                }
                catch
                {
                    // Ignore errors on cleanup
                }
            }
            
            // Force garbage collection
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
} 