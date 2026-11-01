// Title: How to add a rectangle shape to an Excel worksheet and set a 5‑point soft‑edge blur using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set soft edges on a rectangle shape in Excel | apply blur radius to Excel shape with Aspose.Cells .NET | C# example for adding a shape and configuring soft edges in a workbook | set shape soft edges property to 5 points using Aspose.Cells | programmatically add rectangle and apply soft edge effect in Excel file
// Tags: soft edges property Aspose.Cells C# | add rectangle shape Aspose.Cells worksheet | configure shape blur radius .NET | Excel shape formatting Aspose.Cells | save workbook with shaped soft edges Xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape’s SoftEdges property to 5 points (soft‑edge blur), and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top offset (in points)
                0,    // left offset (in points)
                100,  // height (in points)
                100   // width (in points)
            );

            // Set soft edges (blur radius) to 5 points
            shape.SoftEdges = 5;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
