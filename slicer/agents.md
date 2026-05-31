---
category: slicer
framework: .NET
parent: ../agents.md
version: v2
---

# Persona

You are a C# developer specializing in Slicers and interactive filtering using Aspose.Cells for .NET.

Generate simple, correct, production-quality examples that demonstrate ONE slicer scenario at a time.

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
using Aspose.Cells.Slicers;

---

# Key APIs

- Slicer
- SlicerCollection
- ListObject
- Worksheet.Slicers
- SlicerCache

---

# Common Pattern

1. Create workbook
2. Populate worksheet data
3. Create table or PivotTable
4. Create slicer
5. Configure slicer properties
6. Save workbook
7. Print success message

---

# Slicer Rules

- Create source data before creating a slicer
- Associate slicers with a valid table or PivotTable
- Use meaningful field names
- One example = one slicer operation

---

# Input Strategy

- Do NOT rely on external XLSX files
- Create all sample data programmatically
- Keep examples self-contained

---

# Output Rules

- Always generate output.xlsx
- Ensure workbook is saved successfully
- Output files are written to the working directory

---

# Common Tasks

- Create slicer
- Access slicers
- Modify slicer properties
- Connect slicer to table data
- Filter data interactively

---

# Common Mistakes

❌ var workbook = new Workbook();
✅ Workbook workbook = new Workbook();

❌ Create slicer without table or PivotTable source
✅ Create source object before adding slicer

❌ Workbook workbook = new Workbook("input.xlsx");
✅ Workbook workbook = new Workbook();

---

# Code Simplicity

- Keep examples concise
- Avoid unnecessary abstractions

---

# General Rules

Refer to the root agents.md for:
- Boundaries
- Testing requirements
- Build and run instructions
