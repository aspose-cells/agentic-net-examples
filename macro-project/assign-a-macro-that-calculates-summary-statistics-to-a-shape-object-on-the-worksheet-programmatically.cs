// Title: Programmatically add a rectangle shape to the first worksheet and save the workbook as XLSX with Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a rectangle shape at row 1, column 0 with a height of 100 points and a width of 50 points on the first worksheet using Aspose.Cells. | Write C# to save the created workbook to an absolute file path in XLSX format, creating the output folder if it does not exist.
// Common Searches: how to add a rectangle shape to an Excel worksheet using Aspose.Cells C# | Aspose.Cells C# create shape with specific row column dimensions | save Aspose.Cells workbook to a custom directory as .xlsx | ensure output folder exists before saving workbook with Aspose.Cells | C# Aspose.Cells example for adding and positioning shapes
// Tags: add rectangle shape Aspose.Cells C# | shape positioning worksheet Aspose.Cells | define shape dimensions Aspose.Cells | save workbook as xlsx Aspose.Cells | create output directory C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, accesses the first Worksheet, adds a rectangle Shape positioned at row 1, column 0 with a height of 100 points and width of 50 points, ensures the target output folder exists, and saves the workbook as an XLSX file to the current directory.
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

            // Add a rectangle shape and obtain the Shape object
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                0,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                100, // height (in points)
                50   // width (in points)
            );

            // Define output path (ensure absolute path)
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.xlsx");

            // Ensure the output directory exists
            string? directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
