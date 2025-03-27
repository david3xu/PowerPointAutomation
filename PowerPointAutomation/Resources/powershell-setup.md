# PowerPoint Automation Project Setup Guide

This guide provides the PowerShell commands to create the recommended project structure for developing a Knowledge Graph PowerPoint automation application using C# and Office Interop.

## Project Structure Setup Commands

Run these commands in a PowerShell terminal to create the directory structure and placeholder files for the project:

```powershell
# Navigate to the project directory
cd .\PowerPointAutomation\PowerPointAutomation\

# Create the main directory structure
mkdir Models
mkdir Slides
mkdir Utilities
mkdir Resources

# Create the core C# files
New-Item -Path "Program.cs" -ItemType "file" -Force
New-Item -Path "KnowledgeGraphPresentation.cs" -ItemType "file" -Force

# Create slide generator files
New-Item -Path "Slides\TitleSlide.cs" -ItemType "file" -Force
New-Item -Path "Slides\ContentSlide.cs" -ItemType "file" -Force
New-Item -Path "Slides\DiagramSlide.cs" -ItemType "file" -Force
New-Item -Path "Slides\ConclusionSlide.cs" -ItemType "file" -Force

# Create model files
New-Item -Path "Models\SlideContent.cs" -ItemType "file" -Force
New-Item -Path "Models\KnowledgeGraphData.cs" -ItemType "file" -Force

# Create utility files
New-Item -Path "Utilities\ComReleaser.cs" -ItemType "file" -Force
New-Item -Path "Utilities\PresentationStyles.cs" -ItemType "file" -Force
New-Item -Path "Utilities\AnimationHelper.cs" -ItemType "file" -Force

# Placeholder for resource files
New-Item -Path "Resources\placeholder.txt" -ItemType "file" -Force -Value "Place your images and other resource files here."

# Verify the structure
Get-ChildItem -Recurse | Where-Object { !$_.PSIsContainer } | Select-Object FullName
```

## Directory Structure Explanation

The project structure follows a modular organization pattern that separates concerns and improves maintainability:

### Core Components

- **Program.cs**: Entry point of the application
- **KnowledgeGraphPresentation.cs**: Main presentation logic and orchestration

### Models Directory

Contains data structures that represent the content for slides:

- **SlideContent.cs**: Base classes for slide content representation
- **KnowledgeGraphData.cs**: Sample data and structures specific to knowledge graphs

### Slides Directory

Contains specialized classes for generating different slide types:

- **TitleSlide.cs**: Creates formatted title slides
- **ContentSlide.cs**: Generates bullet-point content slides
- **DiagramSlide.cs**: Builds interactive knowledge graph diagrams
- **ConclusionSlide.cs**: Creates conclusion slides with summary content

### Utilities Directory

Contains helper classes for common operations:

- **ComReleaser.cs**: Manages COM object cleanup to prevent memory leaks
- **PresentationStyles.cs**: Centralizes style definitions for consistent branding
- **AnimationHelper.cs**: Provides methods for creating PowerPoint animations

### Resources Directory

Stores static assets used in the presentation:

- Images for diagrams
- Icons or logos
- Any other multimedia content

## Reference Management

After creating the structure, you'll need to add the appropriate references to your project:

1. Right-click on "References" in Solution Explorer
2. Select "Add Reference..."
3. Go to the COM tab
4. Find and add:
   - Microsoft Office XX.X Object Library
   - Microsoft PowerPoint XX.X Object Library
   
Where XX.X corresponds to your installed Office version.

## Implementation Steps

1. Copy the implementation code for each file from the project instruction guide
2. Ensure the namespaces match your project structure
3. Build the solution to verify there are no compilation errors
4. Run the application to test basic functionality
5. Implement advanced features incrementally

## Debugging Tips

- Make PowerPoint visible during development (`pptApp.Visible = MsoTriState.msoTrue`)
- Add console output for tracking progress
- Monitor Task Manager for orphaned PowerPoint processes
- Use try/finally blocks to ensure proper COM cleanup even if exceptions occur

## Common Issues and Solutions

1. **COM Exception: "RPC Server is Unavailable"**
   - Verify PowerPoint installation and permissions

2. **Memory Leaks / PowerPoint Processes Remain**
   - Ensure all COM objects are properly released with `Marshal.ReleaseComObject()`

3. **Shapes Not Appearing as Expected**
   - Check coordinate calculations and units

4. **Animation Not Working**
   - Verify animation sequence and shape references

## Next Steps

After setting up the project structure:

1. Implement the core presentation generation logic
2. Test with basic slides before adding complex features
3. Add knowledge graph diagram capability
4. Implement animations to demonstrate relationships
5. Add error handling and robustness features
6. Test on different PowerPoint versions if needed

The modular structure allows you to build and test components incrementally, focusing on one feature at a time.
