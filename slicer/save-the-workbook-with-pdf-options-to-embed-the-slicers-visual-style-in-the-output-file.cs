using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input_with_slicer.xlsx";
            const string outputPath = "output_with_slicer.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains a slicer
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to preserve slicer visual styles
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}