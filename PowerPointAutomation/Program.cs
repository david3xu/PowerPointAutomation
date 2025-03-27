using System;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using PowerPointAutomation.Utilities;

namespace PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check command line arguments - if "test" is provided, run compatibility tests
            if (args.Length > 0 && args[0].ToLower() == "test")
            {
                RunCompatibilityTests();
                return;
            }

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

        /// <summary>
        /// Runs compatibility tests to verify Office interop works across different versions
        /// </summary>
        static void RunCompatibilityTests()
        {
            Console.WriteLine("Running Office compatibility tests...");
            
            // Define output path for test report
            string reportPath = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop), "CompatibilityTestReport.txt");
                
            // Create test runner and run tests
            var testRunner = new OfficeCompatibilityTest();
            testRunner.RunAllTests(reportPath);
            
            // Open the report file
            try
            {
                Console.WriteLine("Opening the test report...");
                System.Diagnostics.Process.Start(reportPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening report: {ex.Message}");
            }
            
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}