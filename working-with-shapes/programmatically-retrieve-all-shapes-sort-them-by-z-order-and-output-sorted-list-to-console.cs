// Title: C# – Retrieve All Worksheet Shapes and Sort by Z‑Order with Aspose.Cells
// Description: Loads an Excel workbook, accesses the first worksheet, extracts its ShapeCollection, orders the shapes by the ZOrderPosition property, and prints each shape’s name, type, and Z‑order value to the console.
// Keywords: Aspose.Cells shape collection | C# get worksheet shapes | sort shapes by ZOrderPosition | list Excel shapes .NET | shape Z‑order Aspose | enumerate worksheet shapes | Aspose.Cells ordering shapes | C# console output shapes
// Common Searches: Aspose.Cells sort shapes by Z‑order C# | list all shapes in Excel worksheet using Aspose.Cells | retrieve shape collection and ZOrderPosition Aspose | C# code to order Excel shapes by Z‑order | how to get shape Z‑order position with Aspose.Cells
// Developer Intent: Obtain every shape on a worksheet, arrange them by their Z‑order, and display the ordered list.
// Use Cases: Ensure correct visual layering when exporting a sheet to PDF or image formats. | Programmatically adjust shape order before batch formatting or alignment operations. | Validate that important annotations or graphics appear on top in generated reports.
// AI Prompts: Create C# code that moves a selected shape to the front of the Z‑order after retrieving the sorted collection with Aspose.Cells. | Show how to export the sorted shape list (name, type, ZOrderPosition) to a CSV file instead of the console. | Demonstrate grouping shapes by Z‑order ranges and applying distinct formatting to each group using Aspose.Cells.

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, accesses the first worksheet, extracts its ShapeCollection, orders the shapes by the ZOrderPosition property, and prints each shape’s name, type, and Z‑order value to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get all shapes in the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Sort shapes by their Z-order position (ascending)
        var sortedShapes = shapes
            .Cast<Shape>()
            .OrderBy(s => s.ZOrderPosition)
            .ToList();

        // Output the sorted list to the console
        Console.WriteLine("Shapes sorted by Z-order:");
        foreach (var shape in sortedShapes)
        {
            Console.WriteLine($"Name: {shape.Name}, Type: {shape.Type}, ZOrderPosition: {shape.ZOrderPosition}");
        }
    }
}
