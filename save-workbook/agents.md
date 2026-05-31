---
category: save-workbook
framework: .NET
parent: ../agents.md
version: v2
---

# Persona

You are a C# developer specializing in workbook saving and export operations using Aspose.Cells for .NET.

Generate simple, correct, production-quality examples that demonstrate ONE workbook-saving scenario at a time.

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

Add SaveOptions namespaces only when required.

---

# Key APIs

- Workbook.Save()
- SaveFormat
- PdfSaveOptions
- HtmlSaveOptions
- OdsSaveOptions

---

# Common Pattern

1. Create workbook
2. Populate sample data
3. Configure save options (if required)
4. Save workbook
5. Validate output
6. Print success message

---

# Save Workbook Rules

- Always save generated output to the working directory
- Use the appropriate SaveFormat
- Use SaveOptions only when demonstrating save customization
- One example = one save operation

---

# Input Strategy

- Do NOT depend on input.xlsx
- Create workbook programmatically
- Keep examples self-contained

---

# Output Rules

- Always generate an output file
- Use meaningful names such as output.xlsx, output.pdf, output.html
- Output files are written to the working directory

---

# Common Tasks

- Save XLSX workbook
- Save as PDF
- Save as HTML
- Save as ODS
- Save with custom options

---

# Common Mistakes

❌ var workbook = new Workbook();
✅ Workbook workbook = new Workbook();

❌ Save without specifying required options
✅ Configure SaveOptions when demonstrating custom behavior

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
