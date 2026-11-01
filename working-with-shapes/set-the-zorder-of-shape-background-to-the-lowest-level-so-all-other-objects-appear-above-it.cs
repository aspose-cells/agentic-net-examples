// Title: How to move a shape named "Background" to the back layer in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to send background shape to back | C# code to place a specific Excel shape behind all other objects with Aspose.Cells | move Excel shape to backmost layer using Aspose.Cells .NET API | change Z-order of a shape named Background in a workbook with Aspose.Cells | Aspose.Cells example for adjusting shape layering in an Excel file
// Tags: Aspose.Cells set shape ZOrderPosition | C# Aspose.Cells shape ordering | Excel shape back layer Aspose.Cells | background shape Z-order Aspose.Cells | modify shape Z-order in .NET workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, finds the shape named "Background" on the first worksheet, sets its ZOrderPosition to 0 to move it to the backmost layer, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Find the shape named "Background" and send it to the back
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.Name == "Background")
                    {
                        // Set Z-order position to 0 (backmost)
                        shape.ZOrderPosition = 0;
                        break; // Assuming only one shape with this name
                    }
                }

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
