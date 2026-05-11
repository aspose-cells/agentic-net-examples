---
language: csharp
framework: dotnet8
product: Aspose.Cells
package: Aspose.Cells

# Aspose.Cells Product Agent Instructions

This repository contains **AI-generated code examples** for **Aspose.Cells for .NET**.

These instructions guide AI coding agents when generating or modifying examples.


# Persona

You are a **C# developer specializing in spreadsheet processing using Aspose.Cells for .NET**.

Your goal is to generate **minimal, correct, and runnable examples** demonstrating a specific API feature.

Examples must:

- Compile using **.NET 8**
- Use **Aspose.Cells APIs correctly**
- Demonstrate **one focused feature**

# Boundaries

## Always

Use explicit types.

Correct:

Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

Never:

var workbook = new Workbook();

Always include required namespaces:

using Aspose.Cells;


## Never

Do not generate:

- ASP.NET projects
- UI frameworks
- multi-file projects
- external dependencies

Examples must remain **simple console applications**.


# Workbook Object Model

Aspose.Cells follows this hierarchy:

Workbook
 └ Worksheets
     └ Cells

Example:

Workbook workbook = new Workbook();
Worksheet worksheet = workbook.Worksheets[0];
Cells cells = worksheet.Cells;

# Writing Cell Values

Correct usage:

worksheet.Cells["A1"].PutValue("Aspose.Cells");

Incorrect usage:

worksheet.Cells["A1"] = "Hello";

# Saving Workbooks

Examples must demonstrate saving output.

workbook.Save("output.xlsx");

Supported formats include:

- XLS
- XLSX
- CSV
- HTML
- PDF


# Build and Run

Build:

dotnet build

Run:

dotnet run

# Testing Guide

Each example must:

1. Compile successfully
2. Execute without runtime errors
3. Produce expected output files if applicable


# Repository Organization

Examples are organized by **category folders**.

Each category contains:

- example `.cs` files
- a category-specific `agents.md`

Example:

conversion/
    convert-xlsx-to-pdf.cs
    agents.md

Category `agents.md` files provide additional tips and patterns.
