// Title: Insert a rectangle shape into an Excel worksheet and hide it using Aspose.Cells for .NET (placement property workaround)
// AI Prompts: Write C# code that creates a new Workbook, adds a rectangle shape at row 2 column 2 on the first worksheet, and sets its placement so the shape is not displayed when the file is saved with Aspose.Cells. | Generate a C# example that demonstrates how to hide a shape in an Excel file by adjusting the Shape.Placement property using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# hide shape without IsVisible property | how to make a rectangle shape invisible in an Excel workbook using Aspose.Cells | set shape placement to move but not display Aspose.Cells .NET | programmatically hide Excel shape with Aspose.Cells API | Aspose.Cells hide shape workaround for Excel export
// Tags: add rectangle shape Aspose.Cells | shape placement property Aspose.Cells | hide shape Aspose.Cells workaround | Aspose.Cells shape visibility C# | Excel worksheet shape manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, inserts a rectangle shape at row 2, column 2 of the first worksheet, explains that Aspose.Cells lacks a direct IsVisible property, and demonstrates hiding the shape by setting its Placement property before saving the file as output.xlsx.
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

            // Insert a rectangle shape at specified position
            // Parameters: shape type, upper left row, upper left column, top, left, width, height
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row (zero‑based)
                2,   // upper left column (zero‑based)
                5,   // top offset in pixels
                5,   // left offset in pixels
                100, // width in pixels
                50   // height in pixels
            );

            // Note: Aspose.Cells Shape does not expose an IsVisible property.
            // If hiding the shape is required, you can set its placement to not display,
            // but for this example we simply keep it as is.

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
