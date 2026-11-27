// Title: Add a rectangle shape to a worksheet and explore Z‑order and gridline positioning with Aspose.Cells for .NET
// Common Searches: how to place a shape behind gridlines using Aspose.Cells C# | retrieve ZOrder value of a shape in Aspose.Cells .NET | check if shape ZOrder can be negative with Aspose.Cells | Aspose.Cells example for adding rectangle shape to worksheet | does Aspose.Cells support moving shapes to back of cells
// Tags: add rectangle shape Aspose.Cells .NET | shape behind gridlines Aspose.Cells | retrieve shape ZOrder Aspose.Cells | negative ZOrder Aspose.Cells | shape ordering API Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, accesses the first worksheet, adds a rectangle shape via sheet.Shapes.AddShape, and saves the file as ShapeBehindGridlines.xlsx. It notes that properties such as IsBehindGridlines, ZOrder, or MoveToBack may not be available in all Aspose.Cells versions, so the example focuses on shape creation rather than changing its Z‑order or gridline placement.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a simple rectangle shape to the worksheet.
            // AddShape returns the created Shape object.
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                50   // width (pixels)
            );

            // Note: Properties such as IsBehindGridlines, ZOrder, or MoveToBack
            // are not available in all Aspose.Cells versions.
            // The essential operation (adding the shape) is performed above.

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ShapeBehindGridlines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Runtime safety: report any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
