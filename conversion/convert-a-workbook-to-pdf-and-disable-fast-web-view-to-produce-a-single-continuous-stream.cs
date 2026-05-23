using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the existing file
                Workbook workbook = new Workbook(inputPath);

                // Create PDF save options (FastWebView not available in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}