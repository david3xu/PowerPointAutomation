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
```

## 2. Cloning the Repository

```bash
# Clone the repository
git clone https://dev.azure.com/YOUR-ORGANIZATION/YOUR-PROJECT-NAME/_git/PowerPointAutomation
cd PowerPointAutomation
```

## 3. Configuring the Development Environment
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

## 4. Creating the .NET Project

```bash
# Initialize new console app
dotnet new console --output PowerPointAutomation --force

# Add package references
dotnet add PowerPointAutomation package Microsoft.Office.Interop.PowerPoint
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
    <PackageReference Include="Microsoft.Office.Interop.PowerPoint" Version="15.0.0" />
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
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            Application pptApplication = new Application();
            Presentations presentations = pptApplication.Presentations;
            Presentation presentation = presentations.Add(MsoTriState.msoTrue);
            Console.WriteLine("Creating PowerPoint Presentation...");
            try
            {
                string outputPath = Path.Combine(Environment.CurrentDirectory, "KnowledgeGraph.pptx");
                presentation.SaveAs(outputPath);
                presentation.Close();
                pptApplication.Quit();
                Console.WriteLine($"Presentation saved to: {outputPath}");
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

## 7. CI/CD Pipeline with Azure Pipelines

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

## 8. Debugging and Testing

```bash
# Run with debugging
dotnet run --project PowerPointAutomation

# Run specific tests
dotnet test PowerPointAutomation

# View detailed logs
export DOTNET_CLI_LOGGING_LEVEL=detailed
dotnet run --project PowerPointAutomation
```

