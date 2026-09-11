// Title: Add a rectangle shape to an Excel worksheet and position it using pixel coordinates with Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on the first worksheet, assign 50 to its Top and 100 to its Left (pixel units), then save the workbook as ShapePlacement.xlsx. | Check that the shape's Top and Left values equal the requested pixel coordinates and throw an InvalidOperationException if they do not.
// Common Searches: Aspose.Cells C# set shape top and left in pixels | How to position an Excel shape at exact pixel coordinates using Aspose.Cells | Validate rectangle shape placement after setting Top/Left properties in Aspose.Cells .NET
// Tags: add rectangle shape Aspose.Cells | set shape top left pixel offsets Aspose.Cells | shape placement verification Aspose.Cells | pixel based shape positioning Excel Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new workbook, adds a rectangle shape anchored at cell A1, positions it at 100 px from the left edge and 50 px from the top edge, verifies the placement, and saves the file as ShapePlacement.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Desired pixel coordinates for the shape's top-left corner
            int topPixel = 50;   // Y coordinate in pixels
            int leftPixel = 100; // X coordinate in pixels

            // Define shape size in pixels
            int shapeHeight = 80;
            int shapeWidth = 150;

            // Add a rectangle shape anchored at cell A1 (row 0, column 0)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 0, 0, 0, 0, shapeHeight, shapeWidth);

            // Set the shape's offset in pixels
            shape.Top = topPixel;
            shape.Left = leftPixel;

            // Validate that the shape's Top and Left properties match the requested placement
            if (shape.Top != topPixel || shape.Left != leftPixel)
            {
                throw new InvalidOperationException(
                    $"Shape placement error: Expected Top={topPixel}, Left={leftPixel} but got Top={shape.Top}, Left={shape.Left}");
            }

            // Save the workbook to a file
            workbook.Save("ShapePlacement.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
