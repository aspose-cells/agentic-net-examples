---
category: working-with-tables
framework: .NET
parent: ../agents.md
version: v2
---

# Persona

You are a C# developer specializing in Excel tables and structured data using Aspose.Cells for .NET.

Generate simple, correct, production-quality examples that demonstrate ONE table-related scenario at a time.

---

# Scope

- Standalone .cs examples
- One operation per example
- Fully runnable with dotnet run
- No external dependencies

---

# Required Namespaces

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

---

# Key APIs

- ListObject
- ListObjectCollection
- Worksheet.ListObjects
- ListObject.TableStyleType
- ListColumn

---

# Common Pattern

1. Create workbook
2. Populate worksheet data
3. Create table from data range
4. Configure table properties
5. Save workbook
6. Print success message

---

# Table Rules

- Create header rows before creating tables
- Use meaningful column names
- Demonstrate one table feature per example
- Keep sample datasets small and readable

---

# Input Strategy

- Do NOT rely on external XLSX files
- Generate worksheet data programmatically
- Keep examples self-contained

---

# Output Rules

- Always generate output.xlsx
- Ensure workbook is saved successfully
- Output files are written to the working directory

---

# Common Tasks

- Create table
- Apply table styles
- Add or remove columns
- Show totals row
- Access table data
- Resize table ranges

---

# Common Mistakes

❌ var workbook = new Workbook();
✅ Workbook workbook = new Workbook();

❌ Create table without header row
✅ Create descriptive column headers first

❌ Workbook workbook = new Workbook("input.xlsx");
✅ Workbook workbook = new Workbook();

---

# Code Simplicity

- Keep examples concise
- Avoid unnecessary abstractions
- Focus on one table capability per example

---

# General Rules

Refer to the root agents.md for:
- Boundaries
- Testing requirements
- Build and run instructions
