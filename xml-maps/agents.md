---
category: xml-maps
framework: .NET
parent: ../agents.md
version: v2
---

# Persona

You are a C# developer specializing in XML Maps and XML data integration using Aspose.Cells for .NET.

Generate simple, correct, production-quality examples that demonstrate ONE XML Maps scenario at a time.

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

---

# Key APIs

- XmlMap
- XmlMapCollection
- Workbook.Worksheets.XmlMaps
- Cells.ImportXml()
- Cells.ExportXml()

---

# Common Pattern

1. Create workbook
2. Create XML data
3. Import, export, or map XML
4. Validate worksheet content
5. Save workbook
6. Print success message

---

# XML Maps Rules

- Use XmlMapCollection for XML mapping operations
- Keep XML samples small and readable
- Demonstrate one XML feature per example
- Use programmatically generated XML data

---

# Input Strategy

- Do NOT rely on external XML files
- Use inline XML strings
- Keep examples self-contained

---

# Output Rules

- Always generate output.xlsx
- Ensure workbook is saved successfully
- Output files are written to the working directory

---

# Common Tasks

- Import XML
- Export XML
- Create XML maps
- Bind XML data
- Access XML map information

---

# Common Mistakes

❌ var workbook = new Workbook();
✅ Workbook workbook = new Workbook();

❌ Depend on external XML files
✅ Use inline XML samples

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
