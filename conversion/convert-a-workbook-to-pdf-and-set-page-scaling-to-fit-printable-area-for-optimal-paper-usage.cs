// Title: How to convert an Aspose.Cells workbook to PDF in C# and scale pages to fit the printable area
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as PDF while configuring the worksheet PageSetup to fit all columns on one page. | Show a C# example of setting PdfSaveOptions in Aspose.Cells to produce a single‑page‑per‑sheet PDF with all columns forced onto the page. | Provide a C# snippet that populates a workbook, applies SetFitToPages(1,0), and saves it as a print‑optimized PDF.
// Common Searches: Aspose.Cells C# export Excel to PDF with FitToPagesWide set to 1 | How to force all columns onto one PDF page using Aspose.Cells PdfSaveOptions | C# code sample for scaling worksheet to printable area before PDF conversion with Aspose.Cells
// Tags: Aspose.Cells PDF conversion FitToPages | C# worksheet page setup scaling | PdfSaveOptions OnePagePerSheet | force all columns single PDF page | print‑optimized PDF Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // The example creates a workbook, fills it with sample data, configures the worksheet to fit all columns on a single printed page using SetFitToPages(1,0), applies PdfSaveOptions to render each sheet on one page with all columns included, and saves the result as a PDF optimized for printing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure page setup to fit the printable area:
            // Fit all columns on one page (wide = 1) and let the height adjust automatically (tall = 0)
            sheet.PageSetup.SetFitToPages(1, 0);
            // Alternatively you could set the properties directly:
            // sheet.PageSetup.FitToPagesWide = 1;
            // sheet.PageSetup.FitToPagesTall = 0;

            // Create PDF save options to ensure each sheet is rendered on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // All content of a sheet on one page
                AllColumnsInOnePagePerSheet = true,   // Force all columns onto that page
                OptimizationType = PdfOptimizationType.Standard // High print quality
            };

            // Save the workbook as a PDF with the defined options
            string outputPath = "Workbook_FitToPrintableArea.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
    }
}
