---
category: working-with-shapes
framework: .NET
parent: ../agents.md
version: v2
---

# Persona

You are a C# developer specializing in shapes and drawing objects using Aspose.Cells for .NET.

Generate simple, correct, production-quality examples that demonstrate ONE shape-related scenario at a time.

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
using Aspose.Cells.Drawing;

---

# Key APIs

- Shape
- ShapeCollection
- AutoShape
- Worksheet.Shapes
- MsoDrawingType

---

# Common Pattern

1. Create workbook
2. Create worksheet
3. Add shape
4. Configure shape properties
5. Save workbook
6. Print success message

---

# Shape Rules

- Use Worksheet.Shapes for shape operations
- Create visible shapes with meaningful dimensions
- Configure text and formatting when applicable
- One example = one shape operation

---

# Input Strategy

- Do NOT rely on external XLSX files
- Create workbook programmatically
- Keep examples self-contained

---

# Output Rules

- Always generate output.xlsx
- Ensure workbook is saved successfully
- Output files are written to the working directory

---

# Common Tasks

- Add shape
- Modify shape properties
- Add text to shape
- Resize and reposition shape
- Remove shape

---

# Common Mistakes

❌ var workbook = new Workbook();
✅ Workbook workbook = new Workbook();

❌ Create shapes without dimensions
✅ Specify valid row, column, width, and height

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
