// Title: How to send a named "Watermark" shape to the back layer in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move a shape behind worksheet cells | set ZOrderPosition of a named shape in Aspose.Cells workbook | place watermark shape at the back of an Excel sheet with Aspose.Cells | change shape layering order in Aspose.Cells .NET example
// Tags: Aspose.Cells set shape Z-order | C# move Excel shape to back layer | adjust watermark shape layering Aspose.Cells | shape ZOrderPosition property usage | send named shape behind cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, retrieves the shape named "Watermark" from the first worksheet, sets its ZOrderPosition to 0 to move it to the back layer so the underlying cells stay visible, and saves the modified workbook, with error handling for missing files and shape lookup.
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

                // Retrieve the shape named "Watermark"
                Shape watermark = sheet.Shapes["Watermark"];
                if (watermark == null)
                {
                    Console.WriteLine("Shape 'Watermark' not found.");
                }
                else
                {
                    try
                    {
                        // Move the shape to the back layer by setting its Z-order position to the minimum
                        watermark.ZOrderPosition = 0;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to adjust shape order: {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
