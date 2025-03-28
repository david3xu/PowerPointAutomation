# Knowledge Graph PowerPoint Automation

A C# application that automatically generates professional PowerPoint presentations about knowledge graphs using Microsoft Office Interop. The application creates comprehensive slides with custom formatting, interactive diagrams, animations, and speaker notes.

## Features

- Custom master slides with consistent branding
- Interactive knowledge graph diagrams
- Step-by-step animations to demonstrate graph concepts
- Speaker notes for presentation delivery
- Multiple layout types for different content needs
- Memory-optimized COM object management for large presentations
- Individual slide generation for better memory efficiency during testing

## Requirements

- Microsoft PowerPoint (Office 2016 or newer)
- .NET Framework 4.7.2 or .NET 6.0+
- Visual Studio 2019/2022

## Project Structure

```
PowerPointAutomation/
├── Models/
│   ├── KnowledgeGraphData.cs
│   └── SlideContent.cs
├── Slides/
│   ├── TitleSlide.cs
│   ├── IntroductionSlide.cs
│   ├── CoreFeatureSlide.cs
│   ├── ContentSlide.cs
│   ├── ListSlide.cs
│   ├── SummarySlide.cs
│   ├── ComparisonSlide.cs
│   ├── DiagramSlide.cs
│   └── ConclusionSlide.cs
├── Utilities/
│   ├── AnimationHelper.cs
│   ├── ComReleaser.cs
│   ├── IncrementalPresentationGenerator.cs
│   ├── OfficeCompatibility.cs
│   ├── PowerPointOperations.cs
│   ├── PresentationStyles.cs
│   └── ProcessMemoryManager.cs
├── Resources/
│   └── IncreaseProcessMemory.ps1
├── tests/
│   ├── OfficeCompatibilityTest.cs
│   ├── PowerPointTest.cs
│   ├── SimpleTestPresentation.cs
│   ├── run.bat
│   ├── run-simple-test.bat
│   └── Various PowerShell scripts
├── docs/
│   ├── output/
│   ├── PowerPointDebugging.md
│   ├── MemoryOptimizationImprovements.md
│   ├── documentation-overview.md
│   └── Various other documentation files
├── bin/
├── obj/
├── Properties/
├── Program.cs
├── KnowledgeGraphPresentation.cs
├── PowerPointAutomation.csproj
├── App.config
├── packages.config
└── run-full-presentation.bat
```

- **Models/**: Data structures for slide content
  - KnowledgeGraphData.cs: Defines the data model for knowledge graph content
  - SlideContent.cs: General structures for slide content management
- **Slides/**: Specialized slide generators
  - TitleSlide.cs: Implementation of the title slide
  - IntroductionSlide.cs: Implementation of the introduction slide
  - CoreFeatureSlide.cs: Core features presentation slides
  - DiagramSlide.cs: Interactive knowledge graph diagram slides
  - ContentSlide.cs: Base implementation for content slides
  - ListSlide.cs: List-formatted content slides
  - SummarySlide.cs: Summary and recap slides
  - ComparisonSlide.cs: Comparison layout slides
  - ConclusionSlide.cs: Concluding slide implementation
- **Utilities/**: Helper classes for COM interaction and animations
  - AnimationHelper.cs: PowerPoint animation utilities
  - ComReleaser.cs: COM object management for memory optimization
  - IncrementalPresentationGenerator.cs: Multi-stage presentation generation
  - OfficeCompatibility.cs: Office version detection and compatibility 
  - PowerPointOperations.cs: Core PowerPoint operations
  - PresentationStyles.cs: Formatting and style utilities
  - ProcessMemoryManager.cs: Memory optimization utilities
- **Resources/**: Static assets for presentations
  - IncreaseProcessMemory.ps1: System-level memory optimization script
- **tests/**: Test implementations and utilities
  - OfficeCompatibilityTest.cs: Office version compatibility tests
  - PowerPointTest.cs: Basic PowerPoint operations test
  - SimpleTestPresentation.cs: Simplified presentation generation test
  - Various PowerShell and batch files for test automation
- **docs/**: Comprehensive documentation
  - architecture-overview.md: System architecture details
  - PowerPointDebugging.md: Debugging and troubleshooting guide
  - MemoryOptimizationImprovements.md: Memory management strategies
  - user-guide.md: End-user documentation
  - And several other specialized guides for development and operation

## Memory Optimization Features

The application includes sophisticated memory management to avoid COM object leaks and process memory limitations:

- Batch processing of COM objects
- Age-based COM object tracking and release
- Incremental presentation generation mode
- 64-bit process optimization
- System-level memory optimization scripts
- Configurable garbage collection settings
- Individual slide generation for testing and debugging

For details, see [PowerPointDebugging.md](PowerPointAutomation/docs/PowerPointDebugging.md).

## Getting Started

1. Clone this repository
2. Open the solution in Visual Studio
3. Build the solution to restore dependencies
4. For best performance on large presentations:
   - Run the PowerShell script as administrator: `.\PowerPointAutomation\Resources\IncreaseProcessMemory.ps1`
   - Restart your computer to apply system settings
5. Run the application using the optimized batch file: `.\PowerPointAutomation\run-full-presentation.bat`

## Rebuilding the Project

After making changes to the code, you can rebuild the project using MSBuild:

```powershell
# Navigate to the solution directory
cd C:\path\to\PowerPointAutomation

# Rebuild the solution
& "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" PowerPointAutomation.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
```

Alternatively, you can rebuild directly from Visual Studio using the Build menu and selecting "Rebuild Solution".

## Running Options

- **Standard Mode**: Generates the complete presentation in a single process
  ```
  PowerPointAutomation.exe
  ```

- **Incremental Mode**: Generates the presentation in parts for better memory management
  ```
  PowerPointAutomation.exe incremental
  ```

- **Test Mode**: Runs compatibility tests for Office integration
  ```
  PowerPointAutomation.exe test
  ```

- **Single Slide Mode**: Generates only a specific slide for testing
  ```
  PowerPointAutomation.exe slide [slideNumber] [outputPath]
  ```
  Available slide numbers:
  1. Title slide
  2. Introduction slide
  3. Core Components slide
  4. Structural Example slide
  5. Applications slide
  6. Future Directions slide
  7. Conclusion slide

## Generating Individual Slides

For debugging, testing, or reusing specific slides, you can generate just one slide at a time:

### Using the Executable Directly

From the PowerPointAutomation project directory:

```powershell
# Generate Title slide (slide #1)
.\bin\Debug\PowerPointAutomation.exe slide 1 docs\output\KnowledgeGraphPresentation_Slide1.pptx

# Generate slide 2 (Introduction)
.\bin\Debug\PowerPointAutomation.exe slide 2 docs\output\IntroductionSlide.pptx
```

From the solution root directory:

```powershell
# Generate Title slide (slide #1)
.\PowerPointAutomation\bin\Debug\PowerPointAutomation.exe slide 1 .\PowerPointAutomation\docs\output\KnowledgeGraphPresentation_Slide1.pptx
```

### Using the Batch File

When using the batch file, you need to specify an output path:

```powershell
# From the solution root directory:
.\PowerPointAutomation\run-full-presentation.bat slide 2 .\PowerPointAutomation\docs\output\IntroductionSlide.pptx
```

If you don't specify an output path with the batch file, it will use the slide number as the path, which will cause errors.

### Examples for All Slides

These examples can be run from the PowerPointAutomation project directory:

```powershell
# Generate Title slide (slide #1)
.\bin\Debug\PowerPointAutomation.exe slide 1 docs\output\TitleSlide.pptx

# Generate Introduction slide (slide #2)
.\bin\Debug\PowerPointAutomation.exe slide 2 docs\output\IntroductionSlide.pptx

# Generate Core Components slide (slide #3)
.\bin\Debug\PowerPointAutomation.exe slide 3 docs\output\CoreComponentsSlide.pptx

# Generate Structural Example slide (slide #4)
.\bin\Debug\PowerPointAutomation.exe slide 4 docs\output\DiagramSlide.pptx

# Generate Applications slide (slide #5)
.\bin\Debug\PowerPointAutomation.exe slide 5 docs\output\ApplicationsSlide.pptx

# Generate Future Directions slide (slide #6)
.\bin\Debug\PowerPointAutomation.exe slide 6 docs\output\FutureDirectionsSlide.pptx

# Generate Conclusion slide (slide #7)
.\bin\Debug\PowerPointAutomation.exe slide 7 docs\output\ConclusionSlide.pptx
```

**Note:** The output file will have `_Slide#` appended to the filename automatically. For example, if you specify `TitleSlide.pptx`, the actual output will be `TitleSlide_Slide1.pptx`.

This approach offers several advantages:
- Greatly reduced memory usage when working with complex slides
- Faster iteration when making changes to specific slide types
- Easier debugging of memory issues with problematic slides
- Ability to generate just the slides you need for other presentations

## Troubleshooting

If you encounter memory issues:

1. Ensure you're running in 64-bit mode
2. Use the included memory optimization script
3. Try running in incremental mode
4. Generate individual slides to isolate issues
5. See [PowerPointDebugging.md](PowerPointAutomation/docs/PowerPointDebugging.md) for detailed troubleshooting steps

## Documentation

The project includes comprehensive documentation in the `PowerPointAutomation/docs/` directory:

- [documentation-overview.md](PowerPointAutomation/docs/documentation-overview.md): Complete index of all available documentation
- [PowerPointDebugging.md](PowerPointAutomation/docs/PowerPointDebugging.md): Detailed guide for troubleshooting memory and COM-related issues
- [MemoryOptimizationImprovements.md](PowerPointAutomation/docs/MemoryOptimizationImprovements.md): Strategies and techniques used for memory optimization
- [architecture-overview.md](PowerPointAutomation/docs/architecture-overview.md): High-level overview of the system architecture
- [user-guide.md](PowerPointAutomation/docs/user-guide.md): End-user documentation for using the application
- [implementation-summary.md](PowerPointAutomation/docs/implementation-summary.md): Summary of implementation details
- [powerpoint-interop-compat-guide.md](PowerPointAutomation/docs/powerpoint-interop-compat-guide.md): Guide for Office Interop compatibility
- [demo-script.md](PowerPointAutomation/docs/demo-script.md): Script for demonstrating the application
- [devOps-setup.md](PowerPointAutomation/docs/devOps-setup.md): Guide for setting up CI/CD pipelines
- [powershell-setup.md](PowerPointAutomation/docs/powershell-setup.md): PowerShell environment setup guide
- [github-setup-guide.md](PowerPointAutomation/docs/github-setup-guide.md): Git and GitHub workflow setup

The `docs/output/` directory is used to store individually generated slides for testing and demonstration purposes.