using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfPageLimitDemo
{
    // Author: Aspose.Cells .NET example – limit PDF pages using PageCount
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Initialize PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the first page to export (0‑based index). Optional – defaults to first page.
            pdfOptions.PageIndex = 0;

            // Limit the export to a maximum of 3 pages.
            pdfOptions.PageCount = 3;

            // Save the workbook as PDF with the page limitation applied
            workbook.Save("output_limited.pdf", pdfOptions);

            Console.WriteLine("PDF saved with a maximum of 3 pages.");
        }
    }
}