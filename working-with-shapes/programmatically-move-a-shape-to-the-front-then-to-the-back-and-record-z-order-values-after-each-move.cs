// Title: How to change a shape's Z-order (bring to front, send to back) and read its position using Aspose.Cells for .NET
// Common Searches: how to set ZOrderPosition of a shape to bring it to front using Aspose.Cells in C# | C# Aspose.Cells example for sending a shape to the back of the worksheet | retrieve the Z-order index of an Excel shape after reordering with Aspose.Cells | Aspose.Cells .NET code to change shape stacking order in an .xlsx file | read shape.ZOrderPosition after BringToFront call in Aspose.Cells C#
// Tags: Aspose.Cells shape stacking order | set ZOrderPosition property C# | bring shape to front worksheet.Shapes | send shape to back worksheet.Shapes | read shape ZOrderPosition value

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape, prints its initial Z-order, sets ZOrderPosition to 0 to bring the shape to the front, prints the updated value, then sets ZOrderPosition to Shapes.Count‑1 to send it to the back, prints the final value, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, width, height
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1,   // upper left row
                1,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                100, // width (in pixels)
                50   // height (in pixels)
            );

            // Display the initial Z-order position
            Console.WriteLine("Initial Z-order: " + shape.ZOrderPosition);

            // Bring the shape to the front by setting its Z-order to the topmost position (0)
            shape.ZOrderPosition = 0;
            Console.WriteLine("After BringToFront Z-order: " + shape.ZOrderPosition);

            // Send the shape to the back by setting its Z-order to the lowest position
            shape.ZOrderPosition = worksheet.Shapes.Count - 1;
            Console.WriteLine("After SendToBack Z-order: " + shape.ZOrderPosition);

            // Define output file path
            string outputPath = "ShapeZOrderDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
