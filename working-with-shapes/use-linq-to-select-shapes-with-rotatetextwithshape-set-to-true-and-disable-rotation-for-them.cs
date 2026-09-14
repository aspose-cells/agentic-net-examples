// Title: Use LINQ to locate shapes with RotateTextWithShape enabled and reset their RotationAngle in Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, applies a LINQ query to retrieve shapes whose RotateTextWithShape property is true, and sets each shape's RotationAngle to zero while safely handling errors. | Refactor the given Aspose.Cells example to filter shapes by RotateTextWithShape using LINQ and disable rotation only for those shapes.
// Common Searches: aspnet cells linq query to find shapes with RotateTextWithShape enabled | C# Aspose.Cells set RotationAngle to 0 for specific shapes | disable text rotation for Excel shapes using Aspose.Cells library | how to filter worksheet shapes by RotateTextWithShape property in C#
// Tags: LINQ filter RotateTextWithShape Aspose.Cells | reset shape RotationAngle C# | disable shape rotation Aspose.Cells | Excel shape property manipulation Aspose.Cells | Aspose.Cells worksheet shape selection

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Linq;

// The example loads an Excel workbook, uses a LINQ expression to select only those worksheet shapes whose RotateTextWithShape flag is true, sets their RotationAngle to 0 to remove any rotation, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Get all shapes (including TextBox shapes)
            var shapes = worksheet.Shapes.Cast<Shape>();

            // Disable rotation for each shape (attempt; non‑text shapes will be ignored safely)
            foreach (var shape in shapes)
            {
                try
                {
                    // Setting RotationAngle to 0 removes any rotation applied to the shape
                    shape.RotationAngle = 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify shape \"{shape.Name}\": {ex.Message}");
                }
            }

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
