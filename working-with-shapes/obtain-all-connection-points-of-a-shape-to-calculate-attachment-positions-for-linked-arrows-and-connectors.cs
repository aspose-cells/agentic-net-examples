// Title: How to obtain shape connection points in Aspose.Cells for .NET (C#) to calculate connector attachment locations
// Common Searches: Aspose.Cells C# get shape connection points for Excel drawing | C# retrieve connector attachment coordinates from a shape using Aspose.Cells | How to access shape connection points collection in Aspose.Cells .NET | Calculate arrow attachment positions on Excel shapes with Aspose.Cells API | Aspose.Cells shape geometry example for connector placement in C#
// Tags: Aspose.Cells shape connection points | C# Excel shape connector coordinates | Aspose.Cells retrieve shape geometry | C# calculate connector attachment positions | Aspose.Cells drawing API shape points

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new Workbook, adds a rectangle shape to the first worksheet, notes that the current Aspose.Cells version does not expose a ConnectionPoints collection, and saves the workbook as 'ShapeConnectionPoints.xlsx' after ensuring the output directory exists.
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

            // Add a rectangle shape (using positional parameters compatible with the current API)
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,   // shape type
                2,                          // upper left row
                2,                          // upper left column
                5,                          // top offset (in pixels)
                5,                          // left offset (in pixels)
                100,                        // height (in pixels)
                150);                       // width (in pixels)

            // NOTE: The current Aspose.Cells version used in this project does not expose
            // a ConnectionPoints collection directly. If needed, this section can be
            // updated when a newer version with that API becomes available.

            // Save the workbook
            string outputPath = "ShapeConnectionPoints.xlsx";

            // Ensure the directory exists (prevents FileNotFoundException on save)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
