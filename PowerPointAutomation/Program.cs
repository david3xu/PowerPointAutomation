using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define output path - save to desktop for easy access
            string outputPath = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop), "KnowledgeGraphPresentation.pptx");

            Console.WriteLine("Creating Knowledge Graph presentation...");

            // Create presentation generator instance
            var presentationGenerator = new KnowledgeGraphPresentation();

            try
            {
                // Generate the presentation
                presentationGenerator.Generate(outputPath);
                Console.WriteLine($"Presentation successfully created at: {outputPath}");

                // Open the presentation (optional)
                Console.WriteLine("Opening the presentation for review...");
                System.Diagnostics.Process.Start(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating presentation: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}