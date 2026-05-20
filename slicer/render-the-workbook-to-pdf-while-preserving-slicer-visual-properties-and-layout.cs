using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSlicerPdfDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input_with_slicers.xlsx";
                const string outputPath = "output_preserving_slicers.pdf";

                // Ensure the input workbook exists before loading.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook containing slicers.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to preserve document structure (keeps slicers).
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as a PDF.
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully rendered to PDF with slicer layout preserved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}