using System;
using System.IO;
using Aspose.Cells;

namespace ConvertSparklinesToPdfApp
{
    class ConvertSparklinesToPdf
    {
        static void Main()
        {
            const string inputPath = "input_with_sparklines.xlsx";
            const string outputPath = "output_sparklines.pdf";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook that already contains sparkline groups
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as PDF; sparklines are rendered automatically as inline graphics
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}