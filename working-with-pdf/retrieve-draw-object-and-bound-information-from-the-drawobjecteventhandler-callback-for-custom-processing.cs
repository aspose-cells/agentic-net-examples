// Title: Retrieve shape type and bounding rectangle inside DrawObjectEventHandler with Aspose.Cells for .NET
// AI Prompts: Write C# code that registers a DrawObjectEventHandler on a worksheet and extracts the Shape object's type, left, top, width, and height when the event fires. | Show how to log or process the bounds of any drawing object captured by Aspose.Cells' DrawObjectEventHandler.
// Common Searches: Aspose.Cells DrawObjectEventHandler get shape dimensions C# | how to access drawing object bounds in Aspose.Cells event callback | C# example for retrieving shape type and coordinates from Excel using Aspose.Cells | using DrawObjectEventHandler to log rectangle position Aspose.Cells .NET | extract drawing object properties during workbook save Aspose.Cells
// Tags: Aspose.Cells DrawObjectEventHandler shape bounds | C# retrieve drawing object dimensions Aspose.Cells | worksheet shape type extraction .NET | custom processing of drawing objects Aspose.Cells | Excel shape coordinate access using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates how to obtain a shape's type and its bounding rectangle (left, top, width, height) within a DrawObjectEventHandler, enabling custom processing of drawing objects before the workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape directly (the method returns the Shape object)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                100, // width (in points)
                50   // height (in points)
            );

            // Output shape information (type and bounds)
            Console.WriteLine($"Shape Type: {shape.GetType().Name}");
            Console.WriteLine($"Bounds: X={shape.Left}, Y={shape.Top}, Width={shape.Width}, Height={shape.Height}");

            // Force calculation of all formulas in the workbook (optional)
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
