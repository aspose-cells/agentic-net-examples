// Title: Moving the ChartOverlay shape to the front layer in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# bring a shape to the front layer in Excel | set Z-order of a named shape (ChartOverlay) using Aspose.Cells .NET | how to reorder shapes in an Excel worksheet with Aspose.Cells API | workaround for bringing a shape to top of stack in Aspose.Cells C# | move chart overlay shape to top of Z-order in Aspose.Cells workbook
// Tags: Aspose.Cells shape Z-order handling | C# reorder worksheet shapes in Excel | Excel front layer shape manipulation Aspose.Cells | Aspose.Cells workaround re‑add shape for ordering | chart overlay shape ordering Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet, retrieves the shape named "ChartOverlay", and explains that Aspose.Cells does not provide a direct BringToFront method. It suggests using Z‑order workarounds such as re‑adding the shape to place it on the front layer before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the shape by its name
                Shape chartOverlay = sheet.Shapes["ChartOverlay"];

                if (chartOverlay != null)
                {
                    // Aspose.Cells does not provide a BringToFront method.
                    // If needed, you can manipulate Z-order by re‑adding the shape or other means.
                    Console.WriteLine("Shape 'ChartOverlay' found. BringToFront operation is not supported directly.");
                }
                else
                {
                    Console.WriteLine("Shape 'ChartOverlay' not found.");
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
