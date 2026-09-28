// Title: Create a rectangle shape with a tooltip comment in an Excel file using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, insert a rectangle shape at row 3 column 3, set its AlternativeText to a custom message, and save the file as ShapeWithComment.xlsx using Aspose.Cells. | Write C# code that adds a free‑floating shape, assigns a tooltip comment via the AlternativeText property, and exports the worksheet to an .xlsx file.
// Common Searches: Aspose.Cells C# set tooltip for a shape | Add hover comment to Excel shape using .NET | Insert rectangle shape with mouse‑over text in Aspose.Cells | Create Excel workbook with shape tooltip via Aspose.Cells | C# example for shape tooltip in .xlsx file
// Tags: Aspose.Cells rectangle shape API | shape mouse‑over text implementation | C# configure shape line weight | save workbook with embedded shape | shape placement type configuration

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds a free‑floating rectangle shape to the first worksheet, sets its AlternativeText property to act as a tooltip comment, and saves the file as ShapeWithComment.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 120);

            // Set basic shape properties
            shape.Placement = PlacementType.FreeFloating;
            shape.Line.Weight = 1.0; // Set line weight (thickness)

            // Set tooltip (alternative text) for the shape
            shape.AlternativeText = "This is a tooltip comment for the shape.";

            // Determine output file path
            string outputPath = "ShapeWithComment.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
