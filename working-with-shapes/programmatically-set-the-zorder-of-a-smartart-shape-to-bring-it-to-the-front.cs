using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (placeholder for SmartArt) to the worksheet
            // Parameters: shape type, upper left row, upper left column,
            // row offset, column offset, height (points), width (points)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                100, // height
                100  // width
            );

            // Bring the shape to the front of the Z‑order (positive value moves forward)
            shape.ToFrontOrBack(1);

            // Define output file path
            string outputPath = "SmartArtZOrderDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}