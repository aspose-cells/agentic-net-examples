# Aspose.Cells Agentic .NET Examples

Aspose.Cells Agentic .NET Examples is a repository of AI-agent-ready C# examples for creating, analyzing, modifying, converting, formatting, and automating Excel spreadsheets and workbooks using Aspose.Cells for .NET.

This repository contains build-validated, executable C# examples designed for developers, Coding Agents, AI Coding Agents, GitHub Copilot, Semantic Kernel, MCP clients, and modern Agent Frameworks. Every example is automatically generated, compiled, executed, and validated before publication.

## Overview

Aspose.Cells for .NET is a powerful spreadsheet processing API that enables developers and AI agents to create, edit, calculate, convert, secure, and analyze Microsoft Excel files without requiring Microsoft Excel installation.

### Common Use Cases

- Generate Excel spreadsheets programmatically
- Create financial reports and dashboards
- Build pivot tables and charts
- Import and export CSV, JSON, XML, and Excel data
- Automate workbook processing
- Calculate formulas without Microsoft Excel
- Convert Excel files to PDF, HTML, CSV, JSON, and images
- Protect and encrypt workbooks
- Process large spreadsheet datasets

| Metric | Value |
|----------|----------|
| Repository Type | Agentic .NET Examples |
| Language | C# |
| Framework | .NET 10.0+ |
| Product | Aspose.Cells for .NET |
| Validation | Build + Runtime Verified |
| AI Ready | Yes |
| AGENTS.md Support | Yes |

---

## For AI Coding Agents

This repository is structured for direct use by AI Coding Agents, Coding Agents, GitHub Copilot, Semantic Kernel, MCP clients, and Agent Frameworks.

- **[AGENTS.md](./AGENTS.md)** — repository-wide guidance covering API usage, best practices, anti-patterns, and spreadsheet-specific recommendations
- **Per-category AGENTS.md** — focused guidance inside feature folders
- **[index.json](./index.json)** — machine-readable catalog of examples and metadata
- **Build-validated examples** — all examples compile and execute successfully
- **MCP-compatible structure** — optimized for Model Context Protocol workflows

### Supported AI Development Platforms

- GitHub Copilot
- Semantic Kernel
- Claude Desktop
- Cursor
- Continue.dev
- Windsurf
- OpenAI Agents SDK
- Custom Agent Framework implementations

---

## Repository Structure

- calculate-formulas
- cells-data
- comments-and-notes
- conversion
- document-properties
- encryption-and-protection
- format-cells
- globalization-and-localization
- macro-project
- manage-formulas
- manage-workbook
- managing-ranges
- open-workbook
- pivot-table
- queries-and-connections
- rows-and-columns
- save-workbook
- slicer
- smart-markers
- sparkline
- timeline
- workbook-merger
- working-with-charts
- working-with-html
- working-with-images
- working-with-json
- working-with-pdf
- working-with-shapes
- working-with-tables
- working-with-worksheets
- xml-maps

Each folder contains standalone C# examples that can be compiled and executed independently.

---

## Getting Started

### Prerequisites

- .NET SDK (net10.0 or later)
- Aspose.Cells for .NET
- Valid Aspose license (recommended for production use)

### Install Aspose.Cells

```bash
dotnet add package Aspose.Cells
```

### Run an Example

```bash
cd <CategoryFolder>

dotnet new console -o ExampleProject
cd ExampleProject

dotnet add package Aspose.Cells

# Copy example file as Program.cs

dotnet build
dotnet run
```

---

## Spreadsheet Automation Examples

### Load a Workbook

```csharp
using Aspose.Cells;

Workbook workbook = new Workbook("input.xlsx");
```

### Access a Worksheet

```csharp
Worksheet worksheet = workbook.Worksheets[0];
```

### Write Data

```csharp
worksheet.Cells["A1"].PutValue("Hello World");
```

### Calculate Formulas

```csharp
workbook.CalculateFormula();
```

### Save Workbook

```csharp
workbook.Save("output.xlsx");
```

---

## Why Use Aspose.Cells for .NET?

### Excel Automation Without Microsoft Excel

Create, modify, and process Excel files directly from .NET applications.

### Enterprise Spreadsheet Features

- Formulas and calculations
- Charts and graph generation
- Pivot tables
- Conditional formatting
- Smart markers
- Data import/export
- Workbook protection
- PDF conversion
- JSON processing
- XML mapping

### AI-Agent-Friendly Design

Optimized for:

- Retrieval-Augmented Generation (RAG)
- Agentic workflows
- Semantic search
- Code generation
- Automated code review
- GitHub Copilot suggestions

---

## What Can You Build?

Using Aspose.Cells for .NET and these Agentic .NET Examples, developers and AI Coding Agents can:

- Create Excel files in C#
- Read Excel spreadsheets programmatically
- Modify existing workbooks
- Generate financial reports
- Build pivot tables
- Create charts and dashboards
- Export Excel to PDF
- Convert Excel to JSON
- Import CSV and XML data
- Calculate formulas without Microsoft Excel
- Protect and encrypt workbooks
- Process large spreadsheet datasets

These examples are optimized for GitHub Copilot, Semantic Kernel, MCP clients, Coding Agents, and modern Agent Frameworks.

---

## Important Notes

- **Zero-based indexing**: `Worksheets[0]` is the first worksheet.
- **Core object model**: `Workbook`
- **Supported formats**: XLSX, XLS, XLSM, XLSB, ODS, CSV, TSV, JSON, XML, HTML, PDF, Images.
- **Resource management**: Use `using` statements whenever possible.

---

## Agentic .NET Ecosystem

Other Aspose products with agentic, build-validated example repositories:

| Product | Repository | Focus |
|---------|------------|--------|
| Aspose.Words for .NET | https://github.com/aspose-words/agentic-net-examples | Word processing |
| Aspose.Cells for .NET | https://github.com/aspose-cells/agentic-net-examples | Spreadsheet automation |
| Aspose.PDF for .NET | https://github.com/aspose-pdf/agentic-net-examples | PDF processing |
| Aspose.Slides for .NET | https://github.com/aspose-slides/agentic-net-examples | PowerPoint |
| Aspose.HTML for .NET | https://github.com/aspose-html/agentic-net-examples | HTML |
| Aspose.Imaging for .NET | https://github.com/aspose-imaging/agentic-net-examples | Images |
| Aspose.Email for .NET | https://github.com/aspose-email/agentic-net-examples | Email |
| Aspose.BarCode for .NET | https://github.com/aspose-barcode/agentic-net-examples | Barcodes |

---

## Related Resources

### Official Documentation

- [Aspose.Cells for .NET Documentation](https://docs.aspose.com/cells/net/) — Comprehensive guides, tutorials, and feature documentation for spreadsheet processing, Excel automation, and workbook management.
- [API Reference](https://reference.aspose.com/cells/net/) — Complete class, method, property, and namespace reference for Aspose.Cells for .NET.
- [Release Notes](https://releases.aspose.com/cells/net/release-notes/) — Latest product updates, enhancements, bug fixes, and version history.

### Downloads & Packages

- [Aspose.Cells for .NET NuGet Package](https://www.nuget.org/packages/Aspose.Cells/) — Install Aspose.Cells in .NET applications using NuGet.
- [Aspose.Cells Downloads](https://releases.aspose.com/cells/net/) — Download ZIP packages, installers, and product releases.

### Community & Support

- [Aspose.Cells Free Support Forum](https://forum.aspose.com/c/cells/9) — Get help from Aspose engineers and the developer community.
- [Aspose.Cells Blog](https://blog.aspose.com/category/cells/) — Tutorials, code examples, product announcements, and spreadsheet automation guides.
- [GitHub Issues](https://github.com/aspose-cells/agentic-net-examples/issues) — Report issues, suggest improvements, and track repository enhancements.

### Licensing & Purchase

- [Purchase Aspose.Cells](https://purchase.aspose.com/buy) — Explore commercial licensing options for Aspose.Cells for .NET.
- [Temporary License](https://purchase.aspose.com/temporary-license/) — Obtain a free temporary license for full product evaluation.

### Additional Resources

- [Aspose.Cells Product Page](https://products.aspose.com/cells/net/) — Overview of spreadsheet processing features and capabilities.
- [Aspose.Cells Code Examples](https://github.com/aspose-cells) — Open-source examples and developer resources.
- [Aspose.Cells Documentation Home](https://docs.aspose.com/cells/) — Documentation for all supported platforms and languages.

---

## License

All examples use [Aspose.Cells for .NET](https://products.aspose.com/cells/net/) and require a valid license for production use. See [licensing options](https://purchase.aspose.com/buy).

---

*This repository is maintained by automated code generation. For AI-friendly guidance, see [AGENTS.md](./AGENTS.md). Last updated: 2026-06-22*
