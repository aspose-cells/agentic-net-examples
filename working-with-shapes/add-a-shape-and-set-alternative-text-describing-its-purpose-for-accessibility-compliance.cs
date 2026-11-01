// Title: Add a rectangle shape with alternative text for accessibility in an Excel workbook using Aspose.Cells for C#
// Common Searches: how to add a rectangle shape with alt text to an Excel file using Aspose.Cells C# | setting alternative text on Excel shapes for accessibility with Aspose.Cells .NET | example of drawing a rectangle and assigning alt description in Aspose.Cells C# | using Aspose.Cells to create accessible shapes in an Excel workbook programmatically | C# code to add a shape and define its alternative description for screen readers with Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | set shape alt description Aspose.Cells | Excel shape metadata Aspose.Cells | Aspose.Cells drawing API rectangle shape | shape alternative attributes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape on the first worksheet, assigns alternative text for accessibility compliance, and saves the file as ShapeWithAltText.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet and obtain the Shape object directly
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row (zero‑based)
                2,   // upper left column (zero‑based)
                50,  // top offset in points
                50,  // left offset in points
                200, // width in points
                100  // height in points
            );

            // Set alternative text for accessibility compliance
            shape.AlternativeText = "Rectangle highlighting the sales target area.";

            // Define output file path
            string outputPath = "ShapeWithAltText.xlsx";

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
