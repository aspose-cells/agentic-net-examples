// Title: Add a rectangle shape to an Excel worksheet and assign a unique name using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape to an Excel worksheet using Aspose.Cells C# | assign a custom name to a shape in Aspose.Cells .NET | set shape line weight and name property with Aspose.Cells example | reference an inserted shape later by its Name in Aspose.Cells | C# code to insert a shape and set its properties in an Excel file
// Tags: add rectangle shape Aspose.Cells .NET | shape name property Aspose.Cells | set shape line weight Aspose.Cells | worksheet shape insertion C# | reference Excel shape by name Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, inserts a rectangle Shape at a specific cell location, assigns it a unique Name ("MyUniqueShape1"), sets the line weight, and saves the workbook as ShapeDemo.xlsx.
class ShapeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                2,   // Upper left row (zero‑based)
                2,   // Upper left column (zero‑based)
                5,   // Top offset in pixels
                5,   // Left offset in pixels
                100, // Height in pixels
                200  // Width in pixels
            );

            // Assign a unique name to the shape for later reference
            shape.Name = "MyUniqueShape1";

            // Set visual properties that are available in the current API version
            shape.Line.Weight = 2;
            // Note: Line color, fill color, and fill pattern APIs may vary between versions.
            // If needed, they can be set using the appropriate properties supported by the referenced Aspose.Cells version.

            // Determine output path and ensure the directory exists
            string outputPath = "ShapeDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
