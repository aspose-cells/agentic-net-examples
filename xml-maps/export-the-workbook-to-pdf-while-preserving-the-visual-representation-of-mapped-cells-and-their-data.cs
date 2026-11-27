// Title: Export an Excel workbook containing XML maps to PDF while preserving the visual layout using Aspose.Cells for .NET
// Common Searches: C# how to convert an Excel file with XML maps to PDF using Aspose.Cells and keep the original sheet layout | Aspose.Cells PDFSaveOptions preserve worksheet appearance when saving workbook as PDF in .NET | Export Excel workbook to PDF with one page per sheet and PDF/A‑1b compliance using C# | Calculate formulas before exporting Excel to PDF with Aspose.Cells to ensure mapped cell values are displayed | Save Excel workbook as PDF while maintaining visual representation of mapped cells in .NET
// Tags: PdfSaveOptions OnePagePerSheet Aspose.Cells | Aspose.Cells PDF export layout preservation | calculate formulas before PDF conversion Aspose.Cells | PDF/A-1b compliance Aspose.Cells | Excel XML maps PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace ExportWorkbookToPdf
{
    // The example loads an Excel workbook, forces formula recalculation to update mapped cell values, configures PdfSaveOptions with OnePagePerSheet and PDF/A‑1b compliance to retain the original worksheet layout, saves the workbook as a PDF, and disposes resources.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output paths (adjust as needed)
            string inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.xlsx");
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.pdf");

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);

                // Ensure all formulas are calculated so that the visual representation is up‑to‑date
                workbook.CalculateFormula();

                // Configure PDF save options to preserve the visual layout
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Export each worksheet as a single page to keep the original layout
                    OnePagePerSheet = true,
                    // Produce PDF/A‑1b for better compatibility (optional)
                    Compliance = PdfCompliance.PdfA1b
                };

                // Save the workbook as PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully exported to PDF at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Clean up resources
                workbook?.Dispose();
            }
        }
    }
}
