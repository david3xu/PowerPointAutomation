# Setting Up PowerPoint Automation Project on Azure VM with Azure DevOps

## 1. Setting Up Azure DevOps Repository
First, create and configure your Azure DevOps repository:

```bash
# Install Azure CLI if not already available
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash

# Log in to Azure
az login --use-device-code

# Configure Azure DevOps CLI extension
az extension add --name azure-devops

# Set default organization
az devops configure --defaults organization=https://dev.azure.com/YOUR-ORGANIZATION/

# Set default project
az devops configure --defaults project=YOUR-PROJECT-NAME

# Create a new repository
az repos create --name PowerPointAutomation
```

## 2. Configuring the Development Environment
Set up your Azure VM with the necessary tools:

```bash
# Update package sources
sudo apt-get update

# Install .NET SDK
wget https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y apt-transport-https
sudo apt-get install -y dotnet-sdk-6.0

# Verify .NET installation
dotnet --version

# Install Mono for .NET Framework compatibility (if needed)
sudo apt install -y mono-complete

# Install PowerShell Core (optional)
sudo apt-get install -y powershell
```

## 3. Project Directory Structure Setup

```bash
# Create project directory structure
mkdir -p PowerPointAutomation/Models PowerPointAutomation/Slides PowerPointAutomation/Utilities PowerPointAutomation/Resources PowerPointAutomation/Properties

# Create main program files
touch PowerPointAutomation/Program.cs PowerPointAutomation/KnowledgeGraphPresentation.cs

# Create model files
touch PowerPointAutomation/Models/SlideContent.cs PowerPointAutomation/Models/KnowledgeGraphData.cs

# Create slide generator files
touch PowerPointAutomation/Slides/TitleSlide.cs PowerPointAutomation/Slides/ContentSlide.cs PowerPointAutomation/Slides/DiagramSlide.cs PowerPointAutomation/Slides/ConclusionSlide.cs

# Create utility files
touch PowerPointAutomation/Utilities/ComReleaser.cs PowerPointAutomation/Utilities/PresentationStyles.cs PowerPointAutomation/Utilities/AnimationHelper.cs

# Create resource files directory
mkdir -p PowerPointAutomation/Resources/Images
touch PowerPointAutomation/Resources/placeholder.txt

# Create basic configuration files
touch PowerPointAutomation/PowerPointAutomation.csproj
```

## 4. Creating the .NET Project

```bash
# Initialize new console app
dotnet new console --output PowerPointAutomation --force

# Add package references
dotnet add PowerPointAutomation package DocumentFormat.OpenXml
dotnet add PowerPointAutomation package System.Drawing.Common
```

## 5. Project Configuration

```bash
cat > PowerPointAutomation/PowerPointAutomation.csproj << 'EOF'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net6.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="2.16.0" />
    <PackageReference Include="System.Drawing.Common" Version="6.0.0" />
  </ItemGroup>
  <ItemGroup>
    <None Update="Resources\**">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>
</Project>
EOF
```

## 6. Implementation Workflow

```bash
# Create an initial Program.cs file
cat > PowerPointAutomation/Program.cs << 'EOF'
using System;
using System.IO;

namespace PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputPath = Path.Combine(Environment.CurrentDirectory, "KnowledgeGraph.pptx");
            Console.WriteLine("Creating Knowledge Graph presentation...");
            try
            {
                Console.WriteLine("Presentation generation not yet implemented");
                Console.WriteLine($"Presentation would be saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating presentation: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
EOF

# Build the initial project
dotnet build PowerPointAutomation

# Run the project
dotnet run --project PowerPointAutomation
```

## 7. Handling Office Interop on Linux

Since Office Interop is not natively supported on Linux, consider using OpenXML SDK:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

// Create a presentation
using (PresentationDocument presentationDocument =
       PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
{
    PresentationPart presentationPart = presentationDocument.AddPresentationPart();
    presentationPart.Presentation = new Presentation();
    // Add slide parts and content here
}
```

## 8. CI/CD Pipeline with Azure Pipelines

```yaml
trigger:
- main
- feature/*

pool:
  vmImage: 'ubuntu-latest'

variables:
  buildConfiguration: 'Release'

steps:
- task: UseDotNet@2
  inputs:
    packageType: 'sdk'
    version: '6.0.x'

- script: dotnet build --configuration $(buildConfiguration)
  displayName: 'dotnet build $(buildConfiguration)'

- script: dotnet test --configuration $(buildConfiguration) --no-build
  displayName: 'dotnet test $(buildConfiguration)'

- task: DotNetCoreCLI@2
  inputs:
    command: 'publish'
    publishWebProjects: false
    arguments: '--configuration $(buildConfiguration) --output $(Build.ArtifactStagingDirectory)'
    zipAfterPublish: true
  displayName: 'dotnet publish $(buildConfiguration)'

- task: PublishBuildArtifacts@1
  inputs:
    pathtoPublish: '$(Build.ArtifactStagingDirectory)'
    artifactName: 'PowerPointAutomation'
```

## 9. Debugging and Testing

```bash
# Run with debugging
dotnet run --project PowerPointAutomation

# Run specific tests
dotnet test PowerPointAutomation

# View detailed logs
export DOTNET_CLI_LOGGING_LEVEL=detailed
dotnet run --project PowerPointAutomation
```

