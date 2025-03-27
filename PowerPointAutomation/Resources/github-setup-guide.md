# GitHub Repository Setup for PowerPoint Automation Project

This guide provides a methodical approach to creating and initializing a GitHub repository for your Knowledge Graph PowerPoint Automation project using PowerShell commands. Each step is explained with the relevant commands and rationale.

## Prerequisites

Before beginning, ensure you have:

- Git installed on your system
- PowerShell 5.0 or higher
- Basic familiarity with Git concepts
- A GitHub account
- Optional: GitHub CLI installed for streamlined workflow

## Repository Setup Process

### 1. Navigate to Project Directory

First, position yourself in the correct directory to ensure all commands operate on the right files.

```powershell
# Navigate to your project root directory
cd C:\Users\jingu\source\repos\PowerPointAutomation

# Verify your location
Get-Location
```

### 2. Initialize Local Git Repository

Create a new Git repository in your project folder to begin tracking files.

```powershell
# Initialize a new Git repository
git init

# Verify initialization success
git status
```

> **Why this matters**: Initializing a repository creates the hidden `.git` directory that stores all version control information, enabling Git to track changes to your project files.

### 3. Create Essential Repository Files

Every well-structured repository needs standard documentation files to help users and contributors understand the project.

```powershell
# Create a .gitignore file for C# projects
# This prevents unnecessary files from being committed to your repository
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/github/gitignore/master/VisualStudio.gitignore" -OutFile ".gitignore"

# Create a comprehensive README.md file with project description
@"
# Knowledge Graph PowerPoint Automation

A C# application that automatically generates professional PowerPoint presentations about knowledge graphs using Microsoft Office Interop. The application creates comprehensive slides with custom formatting, interactive diagrams, animations, and speaker notes.

## Features

- Custom master slides with consistent branding
- Interactive knowledge graph diagrams
- Step-by-step animations to demonstrate graph concepts
- Speaker notes for presentation delivery
- Multiple layout types for different content needs

## Requirements

- Microsoft PowerPoint (Office 2016 or newer)
- .NET Framework 4.7.2 or .NET 6.0+
- Visual Studio 2019/2022

## Project Structure

- **Models/**: Data structures for slide content
- **Slides/**: Specialized slide generators
- **Utilities/**: Helper classes for COM interaction and animations
- **Resources/**: Static assets for presentations

## Getting Started

1. Clone this repository
2. Open the solution in Visual Studio
3. Build the solution to restore dependencies
4. Run the application to generate a sample presentation
"@ | Out-File -FilePath "README.md" -Encoding utf8

# Create a LICENSE file (MIT License example)
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/licenses/license-templates/master/templates/mit.md" -OutFile "LICENSE.md"
```

> **Why this matters**: The `.gitignore` file prevents unnecessary build artifacts and user-specific files from cluttering your repository. A comprehensive README provides essential information about your project to potential users and contributors.

### 4. Make Initial Commit

Commit all the files you've created to establish the initial repository state.

```powershell
# Check which files will be committed
git status

# Stage all files for commit
git add .

# Make the initial commit with a descriptive message
git commit -m "Initial commit: PowerPoint Automation project structure"

# Verify the commit was created
git log --oneline
```

> **Why this matters**: Making an initial commit establishes the first checkpoint in your project's history. A descriptive commit message helps you and others understand what changes were made.

### 5. Create and Connect to GitHub Repository

#### Option A: Manual GitHub Repository Creation

```powershell
# After manually creating a repository on GitHub through the web interface
# (https://github.com/new), add it as a remote
git remote add origin https://github.com/YOUR-USERNAME/PowerPointAutomation.git

# Verify the remote was added correctly
git remote -v

# Push your local repository to GitHub
git push -u origin master  # Or 'main' depending on your default branch name
```

#### Option B: Using GitHub CLI (More Automated)

If you have GitHub CLI installed, you can create and push to a new repository in one step:

```powershell
# Log in to GitHub CLI if you haven't already
gh auth login

# Create a new GitHub repository (will automatically set it as a remote)
gh repo create PowerPointAutomation --private --source=. --remote=origin --push

# Note: Use --public instead of --private if you want a public repository
```

> **Why this matters**: Connecting your local repository to GitHub allows you to store your code securely in the cloud, enables collaboration with others, and provides access to GitHub's project management features.

## Repository Configuration and Best Practices

### 1. Configure Git Identity

Ensure your commits are properly attributed by setting your identity:

```powershell
# Set your username and email for Git
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"

# Verify the configuration
git config --list | Select-String -Pattern "user"
```

### 2. Set Up Branch Protection (via GitHub Web Interface)

After pushing to GitHub:

1. Go to your repository on GitHub
2. Navigate to Settings > Branches
3. Add a branch protection rule for your main branch
4. Consider enabling:
   - Require pull request reviews before merging
   - Require status checks to pass before merging
   - Include administrators in these restrictions

### 3. Configure Workflow Files

For continuous integration, you can add a GitHub Actions workflow file:

```powershell
# Create .github/workflows directory
mkdir -p .github/workflows

# Create a basic CI workflow for .NET
@"
name: .NET Build and Test

on:
  push:
    branches: [ master, main ]
  pull_request:
    branches: [ master, main ]

jobs:
  build:
    runs-on: windows-latest
    steps:
    - uses: actions/checkout@v3
    - name: Setup .NET
      uses: actions/setup-dotnet@v3
      with:
        dotnet-version: 6.0.x
    - name: Restore dependencies
      run: dotnet restore
    - name: Build
      run: dotnet build --no-restore
    - name: Test
      run: dotnet test --no-build --verbosity normal
"@ | Out-File -FilePath ".github/workflows/dotnet.yml" -Encoding utf8

# Commit and push the workflow file
git add .github
git commit -m "Add GitHub Actions workflow for CI"
git push
```

## Troubleshooting Common Issues

### Authentication Problems

If you encounter authentication issues when pushing to GitHub:

```powershell
# Configure Git to use the credential manager
git config --global credential.helper manager

# For modern Windows systems:
git config --global credential.helper manager-core
```

### Default Branch Name Mismatch

If your default branch is named differently than expected:

```powershell
# Check your current branch
git branch

# If needed, rename your branch
git branch -M main  # Change to the expected name

# Then push with the correct branch name
git push -u origin main
```

### Large Files Rejection

If GitHub rejects pushes containing large files:

```powershell
# First, remove the large file from your last commit
git reset --soft HEAD~1

# Update .gitignore to exclude the large file
Add-Content -Path .gitignore -Value "path/to/large/file.ext"

# Re-commit without the large file
git add .
git commit -m "Initial commit (excluding large files)"
git push -u origin main
```

## Next Steps After Repository Setup

Once your repository is created:

1. **Invite collaborators** (if applicable):
   ```powershell
   gh repo add-collaborator PowerPointAutomation username
   ```

2. **Set up project boards** via the GitHub web interface

3. **Create issue templates** to standardize bug reports and feature requests:
   ```powershell
   mkdir -p .github/ISSUE_TEMPLATE
   # Create templates via the web interface or with PowerShell
   ```

4. **Set up branch-based workflows** for your development process:
   - Create a development branch: `git checkout -b develop`
   - Create feature branches from develop: `git checkout -b feature/new-slide-type develop`

5. **Configure automatic dependency updates** using Dependabot by creating a configuration file:
   ```powershell
   mkdir -p .github
   @"
   version: 2
   updates:
     - package-ecosystem: "nuget"
       directory: "/"
       schedule:
         interval: "weekly"
   "@ | Out-File -FilePath ".github/dependabot.yml" -Encoding utf8
   ```

## Conclusion

You now have a properly configured GitHub repository for your PowerPoint Automation project. This setup follows best practices for .NET projects and provides a solid foundation for version control and collaboration. As your project evolves, you can refine your Git workflow to match your development process and team requirements.

Remember to regularly commit your changes with meaningful commit messages and leverage GitHub's features like Issues and Pull Requests to manage your development process effectively.
