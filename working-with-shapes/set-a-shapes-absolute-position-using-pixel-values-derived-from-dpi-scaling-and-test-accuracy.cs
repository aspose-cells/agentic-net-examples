// Title: Set a rectangle shape's absolute pixel position in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new workbook, adds a rectangle shape, and positions it 200 px from the left edge and 150 px from the top edge with Aspose.Cells. | Demonstrate how to convert shape dimensions from points to pixels using a 96 DPI factor and apply those pixel values when adding the shape via Aspose.Cells. | Write code that prints the shape's UpperLeftRow and UpperLeftColumn after placement and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# place shape at specific pixel offset in worksheet | How to convert points to pixels for shape size in Aspose.Cells | Set absolute position of rectangle shape using pixel coordinates in Excel with Aspose.Cells | Retrieve row and column of shape placed with pixel offsets in Aspose.Cells
// Tags: pixel offset shape positioning Aspose.Cells | convert points to pixels DPI Aspose.Cells | add rectangle shape with absolute coordinates C# | UpperLeftRowOffset column offset Aspose.Cells | save workbook after shape placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, converts shape size from points to pixels using a 96 DPI factor, adds a rectangle shape at pixel offsets (200 px left, 150 px top) on the first worksheet, outputs the shape's row and column indices, and saves the file as ShapePositionTest.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // DPI scaling factor (standard screen DPI) - kept for reference
            const double dpi = 96.0;

            // Desired absolute position in pixels
            int pixelX = 200; // distance from the left edge of the worksheet
            int pixelY = 150; // distance from the top edge of the worksheet

            // Shape size (in points; 1 point = 1/72 inch)
            double shapeWidthPoints = 100; // arbitrary width
            double shapeHeightPoints = 50; // arbitrary height

            // Convert size from points to pixels (approximate conversion)
            int shapeWidthPixels = (int)Math.Round(shapeWidthPoints * dpi / 72.0);
            int shapeHeightPixels = (int)Math.Round(shapeHeightPoints * dpi / 72.0);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape at cell (0,0) with the specified pixel offsets
            // Parameters: shape type, upper left row, upper left column,
            //            row offset (pixels), column offset (pixels),
            //            width (pixels), height (pixels)
            Shape rect = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                0,               // UpperLeftRow
                0,               // UpperLeftColumn
                pixelY,          // UpperLeftRowOffset (pixels)
                pixelX,          // UpperLeftColumnOffset (pixels)
                shapeWidthPixels,
                shapeHeightPixels);

            // Output basic shape placement info (cell indices)
            Console.WriteLine($"Shape placed at row {rect.UpperLeftRow}, column {rect.UpperLeftColumn}.");

            // Save the workbook to verify the result visually if needed
            string outputPath = "ShapePositionTest.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
