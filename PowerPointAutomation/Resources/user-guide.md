# Knowledge Graph PowerPoint Automation - Usage Guide

This guide explains how to build, run, and customize the Knowledge Graph PowerPoint Automation tool. The application automatically generates professional PowerPoint presentations about knowledge graphs with custom formatting, diagrams, animations, and speaker notes.

## Prerequisites

Before running the application, ensure you have the following prerequisites installed:

1. **Visual Studio** (2019 or 2022) with .NET Desktop Development workload
2. **Microsoft PowerPoint** (Office 2016 or newer)
3. **.NET Framework 4.7.2** or newer (comes with Visual Studio)

## Building the Solution

1. **Open the solution in Visual Studio**
   - Double-click the `PowerPointAutomation.sln` file or open it from within Visual Studio

2. **Restore NuGet packages**
   - Right-click on the solution in Solution Explorer and select "Restore NuGet Packages"
   - Alternatively, in the Package Manager Console, run:
     ```
     Update-Package -reinstall
     ```

3. **Build the solution**
   - From the Build menu, select "Build Solution" (or press Ctrl+Shift+B)
   - Verify that the build completes successfully with no errors

## Running the Application

### Method 1: Within Visual Studio

1. Set `PowerPointAutomation` as the startup project (if not already set)
2. Press F5 to start debugging, or Ctrl+F5 to run without debugging

### Method 2: Command Line

1. Navigate to the output directory (usually `bin\Debug` or `bin\Release`)
2. Run the executable:
   ```
   PowerPointAutomation.exe
   ```

### Expected Output

When you run the application:

1. The console window will display progress messages as the presentation is generated
2. PowerPoint will launch in the background (or foreground, depending on visibility settings)
3. The presentation will be automatically created and saved to your desktop as "KnowledgeGraphPresentation.pptx"
4. If configured to do so, the presentation will open automatically after generation

## Customizing the Presentation

### Modifying Content

The presentation content is stored in the `KnowledgeGraphData.cs` file. To modify the content:

1. Open `Models\KnowledgeGraphData.cs`
2. Locate the `GetSamplePresentation()` method 
3. Modify the slide content in the dictionary:
   ```csharp
   slides.Add(2, new BulletSlideContent(
       "Your Custom Title",
       new string[] {
           "Your custom bullet point 1",
           "Your custom bullet point 2",
           // Add more bullet points as needed
       },
       "Your custom speaker notes"
   ));
   ```

### Changing Theme Colors

To change the theme colors:

1. Open `KnowledgeGraphPresentation.cs`
2. Locate the color definitions at the top of the class:
   ```csharp
   private readonly Color primaryColor = Color.FromArgb(31, 73, 125);    // Dark blue
   private readonly Color secondaryColor = Color.FromArgb(68, 114, 196); // Medium blue
   private readonly Color accentColor = Color.FromArgb(237, 125, 49);    // Orange
   ```
3. Modify these colors to your desired values
4. For more extensive theme customization, modify the `ApplyCustomTheme()` method

### Adding New Slide Types

To add a new slide type:

1. Create a new class in the `Slides` directory
2. Implement the necessary logic for generating the slide
3. Add a method in `KnowledgeGraphPresentation.cs` to create your new slide type
4. Call this method from the `Generate()` method

### Customizing Animations

To customize animations:

1. Open the relevant slide generator class (e.g., `DiagramSlide.cs`)
2. Locate the animation code sections
3. Modify the animation effects, timing, and triggers as needed
4. For complex animations, use the `AnimationHelper.cs` utility class

## Troubleshooting

### COM Exceptions

If you encounter COM exceptions:

1. Make sure Microsoft PowerPoint is properly installed
2. Check that you have the correct references to the PowerPoint Interop assemblies
3. Verify that PowerPoint is not in use by another process
4. Ensure that the COM cleanup code is being executed properly

### Memory Leaks / PowerPoint Processes Remain

If PowerPoint processes remain after execution:

1. Use Windows Task Manager to identify lingering processes
2. Ensure all COM objects are properly released using `ComReleaser.cs`
3. Add additional `Marshal.ReleaseComObject()` calls for any COM objects not being properly cleaned up
4. Force garbage collection using `GC.Collect()` and `GC.WaitForPendingFinalizers()`

### Shapes Not Appearing Correctly

If shapes don't appear as expected:

1. Double-check coordinate calculations
2. Verify that the slide layout being used has the expected placeholders
3. Test with simpler shapes first
4. Log the positions and properties of shapes for debugging

## Advanced Customization

### Creating Custom Templates

For more advanced customization, you can create custom PowerPoint templates:

1. Design a template PPTX file with your desired layouts and styles
2. Modify the code to open this template instead of creating a new presentation:
   ```csharp
   presentation = pptApp.Presentations.Open(templatePath, MsoTriState.msoTrue, 
       MsoTriState.msoFalse, MsoTriState.msoTrue);
   ```

### Data-Driven Content

To make the presentation data-driven:

1. Create a data source (JSON, XML, CSV, database)
2. Implement data loading logic in a new utility class
3. Replace the hardcoded content with dynamically loaded content

### Export Options

To add alternative export options:

1. For PDF export, add the following code after generating the presentation:
   ```csharp
   string pdfPath = Path.ChangeExtension(outputPath, ".pdf");
   presentation.ExportAsFixedFormat(pdfPath, PpFixedFormatType.ppFixedFormatTypePDF);
   ```

2. For other formats, consult the PowerPoint Interop documentation

## Next Steps

Once you've mastered the basics, consider these enhancements:

1. Add a configuration file for easy customization without changing code
2. Implement a user interface for interactive content selection
3. Add support for real-time data visualization
4. Create a template system for brand-specific styling
5. Integrate with external data sources for dynamic content