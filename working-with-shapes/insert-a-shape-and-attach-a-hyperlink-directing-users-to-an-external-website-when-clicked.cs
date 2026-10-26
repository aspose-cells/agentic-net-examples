// Title: Add a rectangle shape with an external hyperlink to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape with a hyperlink in an Excel workbook using Aspose.Cells C# | Aspose.Cells C# set hyperlink on a shape in a worksheet | create a clickable shape that opens a URL in Excel with Aspose.Cells | insert shape into Excel and assign an external URL using Aspose.Cells .NET | Aspose.Cells example for linking a shape to a website
// Tags: Aspose.Cells rectangle shape insertion | assign hyperlink to shape in Excel | C# Aspose.Cells shape URL example | Excel worksheet clickable shape using .NET | Aspose.Cells shape formatting and hyperlink

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The sample creates a new workbook, adds a rectangle shape at row 2 column 3 with specific dimensions, assigns the external URL https://www.example.com to the shape's Hyperlink.Address, optionally sets line weight, and saves the file as ShapeWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                2,    // Row index (0‑based)
                3,    // Column index (0‑based)
                0,    // Top offset in points
                0,    // Left offset in points
                80,   // Height in points
                200); // Width in points

            // Assign a hyperlink that points to an external website
            shape.Hyperlink.Address = "https://www.example.com";

            // Optionally set some visual properties (e.g., line weight)
            shape.Line.Weight = 1.0;

            // Save the workbook to a file
            string outputPath = "ShapeWithHyperlink.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
