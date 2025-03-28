# Documentation Overview

This document provides an overview of all the documentation available in the PowerPoint Automation project.

## Core Documentation

| Document | Description |
|----------|-------------|
| [architecture-overview.md](architecture-overview.md) | High-level overview of the system architecture and component relationships |
| [implementation-summary.md](implementation-summary.md) | Summary of key implementation details and design decisions |
| [user-guide.md](user-guide.md) | End-user documentation for using the application |
| [project-instructions.md](project-instructions.md) | Detailed project instructions and requirements |

## Technical Guides

| Document | Description |
|----------|-------------|
| [PowerPointDebugging.md](PowerPointDebugging.md) | Detailed guide for troubleshooting memory and COM-related issues |
| [MemoryOptimizationImprovements.md](MemoryOptimizationImprovements.md) | Strategies and techniques used for memory optimization |
| [powerpoint-interop-compat-guide.md](powerpoint-interop-compat-guide.md) | Comprehensive guide for Office Interop compatibility |

## Development Setup

| Document | Description |
|----------|-------------|
| [devOps-setup.md](devOps-setup.md) | Guide for setting up CI/CD pipelines |
| [powershell-setup.md](powershell-setup.md) | PowerShell environment setup guide |
| [github-setup-guide.md](github-setup-guide.md) | Git and GitHub workflow setup |

## Presentation and Demo

| Document | Description |
|----------|-------------|
| [demo-script.md](demo-script.md) | Script for demonstrating the application |

## Output Directory

The `output/` directory is used to store individually generated slides for testing and demonstration purposes. This directory is where slides are saved when using the single slide generation feature described in the main README.

## Adding New Documentation

When adding new documentation to this project:

1. Place the markdown file in the appropriate subfolder
2. Update this overview document with a link to your new documentation
3. Include a brief description of the document's purpose
4. Update the main README.md if the documentation is relevant to first-time users

## Documentation Standards

- Use Markdown formatting for all documentation
- Include a clear title at the top of each document
- Use section headers to organize content
- Include code samples where appropriate
- Link to other relevant documentation when needed
- Keep documentation up-to-date with code changes 