// Title: Set Z-order for multiple rectangle shapes and confirm their rendering sequence with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to change shape Z-order position | example of layering rectangle shapes in Excel using Aspose.Cells | retrieve ZOrderPosition of shapes after adding them with Aspose.Cells | verify which shape appears on top in an Aspose.Cells generated workbook | set shape Z-order values programmatically in Aspose.Cells .NET
// Tags: Aspose.Cells shape ZOrderPosition | add rectangle shape Aspose.Cells .NET | Excel shape layering Aspose.Cells | C# verify shape rendering order | Aspose.Cells workbook shape ordering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds three rectangle shapes to the first worksheet, assigns each shape a distinct ZOrderPosition (0 = bottom, 1 = middle, 2 = top), prints the Z-order values to the console for verification, and saves the file as ZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add three rectangle shapes with different Z-order values
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape1 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 50, 100);
            shape1.Name = "Shape1";
            shape1.ZOrderPosition = 0; // Bottom layer

            Shape shape2 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 3, 0, 0, 0, 50, 100);
            shape2.Name = "Shape2";
            shape2.ZOrderPosition = 1; // Middle layer

            Shape shape3 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 0, 0, 50, 100);
            shape3.Name = "Shape3";
            shape3.ZOrderPosition = 2; // Top layer

            // Confirm rendering order by outputting Z-order values
            Console.WriteLine("Z-order verification:");
            foreach (Shape s in sheet.Shapes)
            {
                Console.WriteLine($"{s.Name} => ZOrderPosition = {s.ZOrderPosition}");
            }

            // Determine output file path and ensure directory exists
            string outputPath = "ZOrderDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
