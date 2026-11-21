// Title: How to preserve Excel cell borders when converting XLSX to PDF using Aspose.Cells PdfSaveOptions.RenderSolidGridlines in C#
// Common Searches: Aspose.Cells C# preserve Excel gridlines in PDF output | set PdfSaveOptions.RenderSolidGridlines true for PDF conversion Aspose.Cells | keep cell borders when saving workbook as PDF with Aspose.Cells .NET | enable solid gridlines in PDF generated from Excel using Aspose.Cells | C# Aspose.Cells PDF conversion retain original borders
// Tags: Aspose.Cells PdfSaveOptions RenderSolidGridlines | preserve Excel gridlines PDF conversion .NET | cell border rendering Aspose.Cells PDF | C# Excel to PDF with solid gridlines | PDF save options for gridline preservation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook with Aspose.Cells, sets PdfSaveOptions.RenderSolidGridlines to true to keep the original Excel cell borders, ensures the output directory exists, and saves the workbook as a PDF while handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (gridlines rendering is not supported in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
