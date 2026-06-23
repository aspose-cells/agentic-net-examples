using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourceFile = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourceFile);

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set options for high‑quality output
            // Use standard optimization (high print quality) and Flate compression
            pdfSaveOptions.OptimizationType = PdfOptimizationType.Standard;
            pdfSaveOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optional: calculate formulas before saving
            pdfSaveOptions.CalculateFormula = true;

            // Save the workbook as a PDF file with the specified options
            string outputFile = "output.pdf";
            workbook.Save(outputFile, pdfSaveOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
        }
    }
}