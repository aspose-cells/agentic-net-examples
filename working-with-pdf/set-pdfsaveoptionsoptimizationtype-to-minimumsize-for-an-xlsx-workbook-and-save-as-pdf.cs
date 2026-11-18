// Title: How to set PdfSaveOptions.OptimizationType to MinimumSize for XLSX to PDF conversion using Aspose.Cells in C#
// Common Searches: how to configure Aspose.Cells PdfSaveOptions for minimum PDF size in C# | C# code to export an Excel workbook to PDF with MinimumSize optimization using Aspose.Cells | Aspose.Cells example for smallest PDF output when converting XLSX to PDF | setting PDF optimization type to MinimumSize when saving XLSX as PDF with Aspose.Cells
// Tags: PdfSaveOptions OptimizationType MinimumSize | XLSX to PDF conversion Aspose.Cells | reduce PDF file size Aspose.Cells | Aspose.Cells PDF export configuration | Workbook.Save PDF options C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook, creates a PdfSaveOptions object with OptimizationType set to MinimumSize, and saves the workbook as a PDF, producing the smallest possible PDF file using Aspose.Cells for C#.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing XLSX workbook
                Workbook workbook = new Workbook(inputPath);

                // Create PDF save options (default optimization)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF using the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
