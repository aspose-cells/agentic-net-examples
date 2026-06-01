using System;
using Aspose.Cells;

class LimitPdfPages
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the maximum number of pages to render (e.g., first 3 pages)
        pdfOptions.PageCount = 3;

        // Optional: set the starting page index (0‑based). Default is 0.
        pdfOptions.PageIndex = 0;

        // Save the workbook to PDF using the configured options
        workbook.Save("output_limited.pdf", pdfOptions);
    }
}