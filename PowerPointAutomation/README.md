## Generating Individual Slides

To generate a specific slide instead of the full presentation, you can use the following commands:

### Using the Executable Directly

```powershell
# From the PowerPointAutomation project directory:
.\bin\Debug\PowerPointAutomation.exe slide 1 docs\output\TitleSlide.pptx
```

### Examples for Each Slide

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

Generating individual slides uses less memory and allows for faster iteration during development. It's also helpful for debugging memory issues with specific slides. 

## Rebuilding After Code Changes

If you make changes to the code, you'll need to rebuild the project before testing your changes:

### Using MSBuild (Recommended)

```powershell
# From the solution root directory:
& "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" PowerPointAutomation.sln /p:Configuration=Debug /p:Platform="Any CPU" /t:Rebuild
```

### Using Visual Studio

Alternatively, you can rebuild directly from Visual Studio:
1. Open the solution in Visual Studio
2. Select "Build" > "Rebuild Solution" from the menu

After rebuilding, the updated executable will be available at `.\bin\Debug\PowerPointAutomation.exe`. 