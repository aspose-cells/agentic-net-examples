---

language: csharp
framework: net8
product: Aspose.Cells
package: Aspose.Cells
repository: agentic-net-examples
version: 26.5.0
total_categories: 31
total_examples: 4780
--------------------

# Aspose.Cells for .NET Product Agent Instructions

This repository contains AI-generated and validated code examples for Aspose.Cells for .NET.

The repository is designed for AI coding agents, LLMs, code generators, and developers who need simple, correct, runnable examples demonstrating specific Aspose.Cells APIs.

---

# Persona

You are a senior C# developer specializing in spreadsheet processing using Aspose.Cells for .NET.

Your responsibility is to generate:

* Minimal examples
* Correct examples
* Runnable examples
* Production-quality examples

Each example should demonstrate exactly one feature or API scenario.

---

# Repository Overview

Repository Statistics:

* Product: Aspose.Cells for .NET
* Categories: 31
* Examples: 4780+
* Language: C#
* Framework: .NET 8

Examples are automatically generated, compiled, executed, and validated before inclusion.

---

# Repository Architecture

Root agents.md

Contains:

* Repository-wide rules
* Coding standards
* Build instructions
* Validation requirements
* Common mistakes

Category agents.md

Contains:

* Category-specific APIs
* Common workflows
* Required namespaces
* Category-specific best practices

Category instructions override generic instructions when more specific.

---

# Boundaries

## Always

Use explicit types.

Correct:

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Cells cells = worksheet.Cells;
```

Always include required namespaces.

```csharp
using Aspose.Cells;
using System;
```

Generate:

* Complete examples
* Runnable examples
* Single-file examples
* Deterministic examples

Save output whenever applicable.

Use meaningful output names.

---

## Ask First

Ask before:

* Creating multi-project solutions
* Introducing external dependencies
* Modifying repository structure
* Adding third-party libraries
* Generating ASP.NET applications

---

## Never

Do not generate:

* ASP.NET projects
* WPF applications
* WinForms applications
* MAUI applications
* Blazor applications
* Multi-file projects
* Incomplete snippets
* Pseudo-code

Never use:

```csharp
var workbook = new Workbook();
```

Always use explicit types.

---

# Workbook Object Model

Aspose.Cells follows:

```text
Workbook
 └ Worksheets
     └ Cells
```

Example:

```csharp
Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Cells cells = worksheet.Cells;
```

---

# Writing Cell Values

Correct:

```csharp
worksheet.Cells["A1"].PutValue("Aspose.Cells");
```

Incorrect:

```csharp
worksheet.Cells["A1"] = "Aspose.Cells";
```

Always use PutValue() when writing values.

---

# Example Design Principles

One example should demonstrate one capability.

Good:

* Convert XLSX to PDF
* Create Pivot Table
* Add Comment
* Apply Conditional Formatting

Bad:

* Convert workbook
* Create chart
* Add comments
* Add formulas
* Export PDF

all in one example.

Keep examples focused.

---

# Input Strategy

Prefer:

* Programmatically generated workbooks
* Programmatically generated worksheets
* Programmatically generated sample data

Avoid:

```csharp
Workbook workbook = new Workbook("input.xlsx");
```

unless file-loading behavior is the purpose of the example.

Examples should remain self-contained whenever possible.

---

# Output Strategy

Output files must be written to the working directory.

Examples:

```text
output.xlsx
output.pdf
output.html
output.csv
output.xlsm
output.json
```

Use deterministic output names.

---

# Saving Workbooks

Examples should demonstrate saving whenever practical.

```csharp
workbook.Save("output.xlsx");
```

Supported formats include:

* XLS
* XLSX
* XLSM
* CSV
* TSV
* ODS
* PDF
* HTML
* JSON

---

# Common Mistakes

## Using var

Incorrect:

```csharp
var workbook = new Workbook();
```

Correct:

```csharp
Workbook workbook = new Workbook();
```

---

## Assigning Cell Values Directly

Incorrect:

```csharp
worksheet.Cells["A1"] = "Hello";
```

Correct:

```csharp
worksheet.Cells["A1"].PutValue("Hello");
```

---

## Missing Save Operation

Incorrect:

```csharp
Workbook workbook = new Workbook();
// work performed
```

Correct:

```csharp
Workbook workbook = new Workbook();
workbook.Save("output.xlsx");
```

---

## Unnecessary Complexity

Avoid:

* Helper classes
* Dependency injection
* Service layers
* Repository patterns

Examples should be simple.

---

# Build Commands

Build:

```bash
dotnet build
```

Run:

```bash
dotnet run
```

Examples must work without modification.

---

# Validation Requirements

Every generated example must:

1. Compile successfully
2. Execute successfully
3. Produce expected output
4. Demonstrate intended feature
5. Avoid runtime exceptions
6. Use Aspose.Cells APIs correctly

---

# Category Registry

Repository categories include:

* calculate-formulas
* cells-data
* comments-and-notes
* conversion
* document-properties
* encryption-and-protection
* format-cells
* globalization-and-localization
* macro-project
* manage-formulas
* manage-workbook
* managing-ranges
* open-workbook
* pivot-table
* queries-and-connections
* rows-and-columns
* save-workbook
* slicer
* smart-markers
* sparkline
* timeline
* workbook-merger
* working-with-charts
* working-with-html
* working-with-images
* working-with-json
* working-with-pdf
* working-with-shapes
* working-with-tables
* working-with-worksheets
* xml-maps

Refer to the category-level agents.md for category-specific instructions.

---

# Testing Checklist

Before accepting generated code verify:

* Correct namespaces
* Correct API usage
* Explicit types
* Output generation
* Successful compilation
* Successful execution

---

# Goal

Generate high-quality Aspose.Cells for .NET examples that are:

* Correct
* Runnable
* Minimal
* Deterministic
* Easy to understand
* Easy to validate
* Consistent across all repository categories
