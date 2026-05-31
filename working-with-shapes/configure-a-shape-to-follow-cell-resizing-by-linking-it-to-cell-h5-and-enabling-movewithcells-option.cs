using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape.
            // Parameters: upper left row, upper left column,
            // row offset (pixels), column offset (pixels), width (pixels), height (pixels)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

            // Link the shape to cell H5 (A1‑style address)
            shape.LinkedCell = "H5";

            // Make the shape move and resize together with the linked cell
            shape.Placement = PlacementType.MoveAndSize;

            // Define output file path
            string outputPath = "ShapeLinkedToH5.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}