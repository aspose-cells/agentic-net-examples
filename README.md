# Aspose.Cells for .NET Examples

AI-friendly repository containing validated C# examples for Aspose.Cells for .NET API.

## Overview

This repository provides working code examples demonstrating Aspose.Cells for .NET capabilities. All examples are automatically generated, compiled, and validated using the Aspose.Cells Examples Generator.

## Repository Structure

Examples are organized by feature category:

- `cells-data/`
- `comments-and-notes/`
- `conversion/`
- `document-properties/`
- `encryption-and-protection/`
- `globalization-and-localization/`
- `macro-project/`
- `manage-formulas/`
- `manage-workbook/`
- `managing-ranges/`
- `open-workbook/`
- `queries-and-connections/`
- `save-workbook/`
- `slicer/`
- `smart-markers/`
- `sparkline/`
- `workbook-merger/`
- `working-with-html/`
- `working-with-images/`
- `working-with-json/`
- `working-with-pdf/`
- `working-with-tables/`
- `working-with-worksheets/`
- `xml-maps/`

Each category contains standalone `.cs` files that can be compiled and run independently.

## Getting Started

### Prerequisites

- .NET SDK (net10.0 or compatible version)
- Aspose.Cells for .NET NuGet package
- Valid Aspose license (for production use)

### Running Examples

Each example is a self-contained C# file. To run an example:

cd <CategoryFolder>
dotnet new console -o ExampleProject
cd ExampleProject
dotnet add package Aspose.Cells
# Copy the example .cs file as Program.cs
dotnet build

dotnet run

## Code Patterns

### Loading a Workbook

using (Workbook workbook = new Workbook("input.xlsx"))
{
    // Work with workbook
}

### Accessing Worksheets and Cells

Worksheet worksheet = workbook.Worksheets[0];
Cell cell = worksheet.Cells["A1"];
cell.PutValue("Hello World");

### Saving a Workbook

workbook.Save("output.xlsx");

### Important Notes

- Zero-based indexing: Worksheets use 0-based indexing (Worksheets[0] = first worksheet)
- Core object: Aspose.Cells works with Workbook instead of Document
- Deterministic cleanup: Use using statements where applicable

## Contributing

Examples in this repository are automatically generated. To suggest new examples:

1. Submit tasks to the Aspose.Cells Examples Generator
2. Generated examples are validated via compilation
3. Passing examples are included in repository updates

## Related Resources

- [Aspose.Cells for .NET Documentation](https://docs.aspose.com/cells/net/)  
- [API Reference](https://reference.aspose.com/cells/net/)  
- [Aspose Forum](https://forum.aspose.com/c/cells/9)  
- [AI Agent Guide (agents.md)](./agents.md)  

## License

All examples use Aspose.Cells for .NET and require a valid license for production use. See licensing page on Aspose website.

---

This repository is maintained by automated code generation. For AI-friendly guidance, see agents.md.
