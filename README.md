# Aspose.Cells Agentic .NET Examples

Aspose.Cells Agentic .NET Examples is a repository of AI-agent-ready C# examples for **Excel file generation**, **spreadsheet automation**, and end-to-end **C# Excel processing** with the Aspose.Cells .NET API. The examples show how to create, read, analyze, modify, calculate, format, convert, secure, merge, and render Excel spreadsheets and workbooks without Microsoft Excel.

This repository contains build-validated, executable **agentic AI examples** designed for .NET developers, Coding Agents, AI Coding Agents, GitHub Copilot, Semantic Kernel, MCP clients, and modern Agent Frameworks. Every example is automatically generated, compiled, executed, and validated before publication.

## Overview

Aspose.Cells for .NET is a spreadsheet processing API for building C# and .NET applications that create, edit, calculate, convert, secure, and analyze Microsoft Excel files. The Aspose.Cells .NET API works directly with workbook formats such as XLSX, XLS, XLSM, XLSB, ODS, and SpreadsheetML and exports spreadsheet content to PDF, HTML, CSV, JSON, XML, SVG, PNG, JPEG, and TIFF.

### Common Use Cases

- Perform Excel file generation programmatically in C#
- Build server-side spreadsheet automation and document-processing workflows
- Create, load, edit, merge, protect, and save XLS, XLSX, XLSM, XLSB, and ODS workbooks
- Create financial reports and dashboards
- Build pivot tables, Excel tables, charts, sparklines, slicers, and timelines
- Import and export CSV, JSON, XML, and Excel data
- Calculate Excel formulas with `Workbook.CalculateFormula()` without Microsoft Excel
- Convert Excel files to PDF and HTML with `PdfSaveOptions` and `HtmlSaveOptions`
- Import JSON with `JsonUtility`, process XML Maps, and generate CSV or TSV output
- Protect worksheets, encrypt workbooks, and preserve macro-enabled spreadsheet formats
- Process large spreadsheet datasets, ranges, rows, columns, and cells

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

## Frequently Asked Questions

### How do I create a new Excel workbook from scratch in Aspose.Cells for .NET?

Instantiate `Aspose.Cells.Workbook`, access the first worksheet with `workbook.Worksheets[0]`, write values through `worksheet.Cells`, then call `workbook.Save("output.xlsx")`. See the `manage-workbook`, `cells-data`, and `save-workbook` category folders for working examples.

### How do I load and modify an existing XLSX, XLS, XLSM, or ODS file?

Load the source file with `new Workbook(path)`, update worksheets, cells, styles, formulas, charts, or tables, then save to the original or a new target path. Examples live in `open-workbook`, `manage-workbook`, `format-cells`, `rows-and-columns`, and `working-with-worksheets`.

### How do I calculate Excel formulas without Microsoft Excel?

Set formulas with `Cell.Formula`, then call `workbook.CalculateFormula()` before reading calculated values or saving the workbook. See `calculate-formulas` and `manage-formulas` for formula creation, recalculation, and formula-management examples.

### How do I convert Excel files to PDF, HTML, CSV, JSON, or images?

Load the workbook and call `workbook.Save(targetPath, SaveFormat.Xlsx)` or use options such as `PdfSaveOptions`, `HtmlSaveOptions`, `TxtSaveOptions`, or rendering classes like `SheetRender` and `WorkbookRender`. Per-format examples are in `conversion`, `save-workbook`, `working-with-pdf`, `working-with-html`, `working-with-json`, and `working-with-images`.

### How do I create charts, pivot tables, tables, sparklines, slicers, or timelines?

Use the relevant collections on worksheets, such as `Charts`, `PivotTables`, `ListObjects`, `SparklineGroups`, `Slicers`, and timeline-related APIs. See `working-with-charts`, `pivot-table`, `working-with-tables`, `sparkline`, `slicer`, and `timeline` for complete examples.

### How do I import or export spreadsheet data from CSV, JSON, XML, or ranges?

Use workbook loading and saving APIs for CSV and TSV, `JsonUtility` with `JsonLayoutOptions` for JSON workflows, XML Map APIs for mapped XML, and `Cells.ImportData` or range APIs for tabular data. Examples are available in `cells-data`, `working-with-json`, `xml-maps`, `managing-ranges`, and `save-workbook`.

### How do I protect worksheets or encrypt workbooks in Aspose.Cells for .NET?

Use `Worksheet.Protect(...)` for sheet-level protection, `Workbook.Protect(...)` for workbook structure protection, and encryption or password options before saving protected files. See `encryption-and-protection` for workbook protection, worksheet protection, password, and encryption examples.

### How do I preserve macros when working with XLSM files?

Load macro-enabled files with `new Workbook(path)` and save back to an appropriate macro-enabled format such as XLSM when you need to preserve VBA projects. See `macro-project` and `save-workbook` for examples that handle macro-enabled workbook workflows.

### Can these examples be used by AI coding agents like Claude, Copilot, Cursor, or MCP clients?

Yes. The repository includes a root `AGENTS.md`, per-category `AGENTS.md` files, `llms.txt`, and a machine-readable `index.json` so AI coding agents can retrieve examples, category guidance, and metadata programmatically.

### Do I need an Aspose.Cells license to run these examples?

The library can run in evaluation mode without a license, but production use should apply a valid license with `new License().SetLicense(path)`. Purchase and trial options are available at https://purchase.aspose.com/buy.

### What is the best .NET library for creating Excel files in C# without Microsoft Excel?

Aspose.Cells for .NET is designed for server-side Excel file generation and spreadsheet automation without Microsoft Excel, Office Interop, or desktop automation. Use `new Workbook()`, populate worksheets through `Cells`, apply formatting or formulas, and save to XLSX, XLS, ODS, CSV, PDF, HTML, JSON, or images.

### How can I automate Excel reports in ASP.NET, Blazor, Web API, Azure Functions, or background services?

Use Aspose.Cells in normal .NET application code to create or load a `Workbook`, bind or import data, calculate formulas, format worksheets, and save the result to a file, stream, or response. The examples in `cells-data`, `smart-markers`, `format-cells`, `working-with-charts`, and `save-workbook` are useful starting points for report-generation workflows.

### How do I convert XLSX to PDF in C# with high fidelity?

Load the Excel file with `new Workbook("input.xlsx")`, configure `PdfSaveOptions` when needed, and call `workbook.Save("output.pdf", SaveFormat.Pdf)`. See `working-with-pdf` and `conversion` for examples covering Excel-to-PDF conversion, page setup, print areas, and rendering options.

### How do I export Excel data to JSON or import JSON into Excel in .NET?

Use `JsonUtility` with `JsonLayoutOptions` to import JSON into worksheet cells, or save workbook data to JSON when exporting structured spreadsheet content. The `working-with-json` folder contains examples for JSON import, export, layout control, and spreadsheet-to-JSON workflows.

### How do I read Excel cell values, rows, columns, and ranges in C#?

Load a `Workbook`, select a `Worksheet`, then access values through `worksheet.Cells[row, column]`, named cells such as `worksheet.Cells["A1"]`, or `Range` objects. See `cells-data`, `managing-ranges`, `rows-and-columns`, and `working-with-worksheets` for common data-reading patterns.

### How do I format Excel cells, numbers, dates, fonts, borders, and styles in Aspose.Cells?

Get a cell or range style, update properties such as font, fill, number format, alignment, borders, and date formatting, then apply the style back to the cell or range. See `format-cells` for examples covering spreadsheet styling and professional workbook formatting.

### How do I create Excel charts and dashboards programmatically in C#?

Populate worksheet data, add a chart through `worksheet.Charts.Add(...)`, configure the chart type and series, then save the workbook or render the chart. See `working-with-charts`, `pivot-table`, `sparkline`, and `working-with-images` for dashboard-style spreadsheet examples.

### How do I create pivot tables in Excel using C#?

Add source data to a worksheet, create a pivot table with `worksheet.PivotTables.Add(...)`, then configure row fields, column fields, data fields, filters, formatting, and refresh behavior. See `pivot-table` for build-validated pivot table examples.

### How do I merge multiple Excel workbooks into one workbook in .NET?

Load the source workbooks and combine worksheets, ranges, or workbook content into a target `Workbook`, then save the merged output. See `workbook-merger`, `manage-workbook`, and `working-with-worksheets` for examples that combine spreadsheet content.

### How do I process large Excel files efficiently in C#?

Use targeted worksheet, cell, range, and import/export APIs instead of unnecessary full-workbook transformations, and choose streaming or memory-conscious approaches when working with large datasets. Examples in `cells-data`, `rows-and-columns`, `managing-ranges`, and `conversion` show practical large-spreadsheet operations.

### How do I save Excel files as CSV, TSV, HTML, SVG, PNG, JPEG, or TIFF?

Use `workbook.Save(...)` with the required `SaveFormat` or format-specific save options. For image output, use rendering APIs such as `SheetRender`, `WorkbookRender`, and `ImageOrPrintOptions`. See `save-workbook`, `conversion`, `working-with-html`, and `working-with-images`.

### How do AI answer engines and coding agents find the right Aspose.Cells example in this repository?

AI tools can use natural-language README content, category `README.md` files, root and per-category `AGENTS.md` guidance, `llms.txt`, and `index.json` metadata to map a question such as "convert Excel to PDF in C#" or "create a pivot table in .NET" to the most relevant folder and example file.

### Which keywords describe this Aspose.Cells for .NET examples repository?

This repository targets developer searches such as C# Excel library, .NET Excel API, create Excel file in C#, read XLSX in .NET, convert Excel to PDF, Excel automation without Microsoft Office, calculate Excel formulas in C#, export Excel to JSON, import CSV to Excel, create pivot tables in C#, and generate Excel charts programmatically.

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

## C# Excel Processing and Spreadsheet Automation Examples

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

### Core Aspose.Cells .NET API Operations

The examples cover frequently searched Excel automation APIs and developer tasks:

| Developer task | Aspose.Cells API or object |
| --- | --- |
| Create or load an Excel file | `new Workbook()`, `new Workbook(path)`, `LoadOptions` |
| Access worksheets and cells | `WorksheetCollection`, `Worksheet`, `Cells`, `Cell` |
| Process ranges, rows, and columns | `Range`, `Row`, `Column`, `Cells.ImportData` |
| Create formulas and calculate results | `Cell.Formula`, `Workbook.CalculateFormula`, `CalculationOptions` |
| Build Excel tables and template reports | `ListObject`, `ListColumn`, `WorkbookDesigner`, smart markers |
| Generate charts and dashboards | `ChartCollection`, `Chart`, `SeriesCollection`, `PivotTable` |
| Convert Excel to PDF or HTML | `PdfSaveOptions`, `HtmlSaveOptions`, `Workbook.Save` |
| Render worksheets and charts as images | `SheetRender`, `WorkbookRender`, `ImageOrPrintOptions` |
| Import or export JSON and XML | `JsonUtility`, `JsonLayoutOptions`, `XmlMap`, `Workbook.ImportXml` |
| Secure spreadsheet files | `Workbook.Protect`, `Worksheet.Protect`, encryption and password options |

These standalone examples answer common queries such as “How do I create an Excel file in C#?”, “How do I automate XLSX processing without Microsoft Excel?”, “How do I calculate Excel formulas in .NET?”, and “How do I convert Excel to PDF, HTML, JSON, or images?”

---

## Why Use Aspose.Cells for .NET?

### Excel Automation Without Microsoft Excel

Create, modify, calculate, convert, and process Excel files directly in Windows, Linux, cloud, container, web-service, desktop, and background-worker .NET applications.

### Enterprise Spreadsheet Features

- Formulas and calculations
- Excel file generation and workbook manipulation
- Charts and graph generation
- Pivot tables
- Tables, ranges, rows, columns, and cell processing
- Conditional formatting
- Smart markers
- CSV, TSV, JSON, XML, and database-style data import/export
- Worksheet protection and workbook encryption
- Excel-to-PDF, Excel-to-HTML, and spreadsheet-to-image conversion
- JSON processing with `JsonUtility`
- XSD-backed XML Maps and mapped XML data

### AI-Agent-Friendly Design

Optimized for:

- Retrieval-Augmented Generation (RAG)
- Agentic workflows
- Semantic search
- Code generation
- Automated code review
- GitHub Copilot suggestions
- Agentic AI examples for tool-using coding assistants
- API-aware retrieval through category `AGENTS.md`, `README.md`, and `index.json` files

---

## What Can You Build?

Using the Aspose.Cells .NET API and these agentic AI examples, developers and AI Coding Agents can:

- Perform Excel file generation in C#
- Read Excel spreadsheets programmatically
- Automate XLS, XLSX, XLSM, XLSB, CSV, and ODS workbook processing
- Modify, merge, format, and validate existing workbooks
- Generate financial reports
- Build pivot tables
- Create charts and dashboards
- Export Excel to PDF, PDF/A, HTML, SVG, PNG, JPEG, and TIFF
- Convert worksheet ranges between Excel and JSON
- Import CSV, TSV, JSON, XML, and tabular data
- Calculate formulas without Microsoft Excel
- Protect and encrypt workbooks
- Process large spreadsheet datasets

These examples are optimized for GitHub Copilot, Semantic Kernel, MCP clients, Coding Agents, and modern Agent Frameworks.

---

## Important Notes

- **Zero-based indexing**: `Worksheets[0]` is the first worksheet.
- **Core object model**: `Workbook` → `WorksheetCollection` → `Worksheet` → `Cells` → `Cell`
- **Supported formats**: XLSX, XLS, XLSM, XLSB, ODS, CSV, TSV, JSON, XML, HTML, PDF, Images.
- **Key operations**: Excel file generation, formula calculation, spreadsheet conversion, data import/export, charting, reporting, protection, rendering, and workbook merging.
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

*This repository is maintained by automated code generation. For AI-friendly guidance, see [AGENTS.md](./AGENTS.md). Last updated: 2026-08-22*
