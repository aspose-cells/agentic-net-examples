// Title: Add a rectangle shape to an Excel worksheet with Aspose.Cells for .NET and set a negative Z‑order to position it behind gridlines
// AI Prompts: Generate C# code that inserts a rectangle shape at a specific cell using Aspose.Cells, sets ZOrderPosition = -1, prints the Z-order value, and saves the workbook. | Demonstrate how to confirm that a shape’s ZOrderPosition places it beneath the worksheet gridlines in an Aspose.Cells .NET workbook.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition negative to hide behind gridlines | how to place a rectangle shape under Excel gridlines using Aspose.Cells for .NET | C# example of moving a shape behind worksheet cells with Aspose.Cells | check if a shape is rendered below gridlines in an Aspose.Cells workbook | Aspose.Cells shape layering order documentation for .NET
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells ZOrderPosition usage | shape behind gridlines Aspose.Cells | C# Excel shape Z-order | validate shape layering Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample creates a new workbook, adds a rectangle shape at row 2 column 2, assigns ZOrderPosition = -1 to place the shape behind the gridlines, outputs the Z-order value for verification, and saves the file as ShapeBehindGridlines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                100, // height (in points)
                50   // width (in points)
            );

            // Set Z-order to a negative value to place it behind other objects (including gridlines)
            shape.ZOrderPosition = -1;

            // Output shape property for verification
            Console.WriteLine($"Shape ZOrderPosition: {shape.ZOrderPosition}");

            // Define output file path
            string outputPath = "ShapeBehindGridlines.xlsx";

            // Save the workbook (optional visual verification)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
