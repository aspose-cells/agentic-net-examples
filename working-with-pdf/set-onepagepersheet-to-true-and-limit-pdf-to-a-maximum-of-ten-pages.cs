// Title: Export Excel to PDF with OnePagePerSheet and a 10‑page limit using Aspose.Cells (C#)
// Description: C# example that creates a workbook, fills it with data, and saves it as a PDF with PdfSaveOptions configured to render each worksheet on a single page (OnePagePerSheet = true) and stop after ten pages (PageCount = 10).
// Keywords: Aspose.Cells PDF export | OnePagePerSheet true | limit PDF pages | PdfSaveOptions PageCount | C# Excel to PDF | Aspose.Cells pagination
// Common Searches: Aspose.Cells set OnePagePerSheet in PDF | How to limit PDF page count with Aspose.Cells | C# export Excel to PDF with max pages | PdfSaveOptions PageCount example
// Developer Intent: Generate a PDF from an Excel workbook where each sheet fits on one page and the output contains no more than ten pages.
// Use Cases: Produce a printable summary where each worksheet occupies a single page. | Create a quick preview PDF that includes only the first ten pages of a large workbook. | Control file size by capping the PDF page count while preserving a one‑page‑per‑sheet layout.
// AI Prompts: Write C# code that sets OnePagePerSheet = true and limits the PDF to 10 pages with Aspose.Cells. | Explain the interaction between OnePagePerSheet and PdfSaveOptions.PageCount in Aspose.Cells. | Show alternative methods to truncate an Aspose.Cells‑generated PDF after a specific number of pages.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // C# example that creates a workbook, fills it with data, and saves it as a PDF with PdfSaveOptions configured to render each worksheet on a single page (OnePagePerSheet = true) and stop after ten pages (PageCount = 10).
    class Program
    {
        static void Main()
        {
            // Create a new workbook (you can also load an existing file using new Workbook("input.xlsx"))
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate pagination
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet is rendered on a single page
                OnePagePerSheet = true,

                // Limit the output to a maximum of 10 pages
                PageCount = 10
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
