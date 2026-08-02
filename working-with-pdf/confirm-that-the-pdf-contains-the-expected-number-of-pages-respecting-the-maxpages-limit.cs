// Title: Limit PDF Output Pages with Aspose.Cells for .NET Using PageSavingCallback
// Description: Demonstrates how to restrict the number of pages saved when exporting an Aspose.Cells workbook to PDF. The sample creates a workbook, fills 500 rows, defines a print area, sets PdfSaveOptions.PageCount, and attaches an IPageSavingCallback that logs each page and stops the export once the configured maximum (e.g., 3 pages) is reached.
// Keywords: Aspose.Cells PDF page limit | PdfSaveOptions PageCount C# | IPageSavingCallback Aspose.Cells | limit PDF pages Aspose | C# truncate PDF export | Aspose.Cells pagination control | PDF export performance Aspose | Aspose.Cells page count callback | C# limit exported PDF pages | Aspose.Cells PDF preview
// Common Searches: How to limit number of pages when saving workbook to PDF with Aspose.Cells | Aspose.Cells IPageSavingCallback example for page count | Retrieve total page count during PDF export in Aspose.Cells | Stop PDF generation after N pages using Aspose.Cells | Set maximum pages in PdfSaveOptions Aspose.Cells
// Developer Intent: Prevent the generated PDF from exceeding a predefined maximum page count.
// Use Cases: Create a preview PDF that contains only the first few pages of a large workbook. | Enforce document size limits in automated reporting pipelines. | Improve export performance by aborting PDF generation after the required pages are saved. | Validate workbook pagination before sending the file to downstream systems. | Log page‑by‑page export progress for audit or debugging purposes.
// AI Prompts: Show a C# test that opens the saved PDF with Aspose.Pdf and asserts it contains exactly 3 pages. | Explain how to modify PageCountCallback to throw a custom exception when the workbook exceeds the allowed page limit. | Demonstrate retrieving the final page count after saving without using a callback, using Aspose.Cells APIs. | Provide guidance on configuring PdfSaveOptions to embed page numbers only for the exported pages. | Suggest ways to combine this callback with a progress bar in a WinForms application.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPageCountDemo
{
    // Demonstrates how to restrict the number of pages saved when exporting an Aspose.Cells workbook to PDF. The sample creates a workbook, fills 500 rows, defines a print area, sets PdfSaveOptions.PageCount, and attaches an IPageSavingCallback that logs each page and stops the export once the configured maximum (e.g., 3 pages) is reached.
    class Program
    {
        static void Main()
        {
            // Define the maximum number of pages that should be saved to the PDF
            int maxPages = 3;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with enough rows to generate multiple pages
            for (int i = 0; i < 500; i++)
            {
                worksheet.Cells[i, 0].Value = $"Row {i + 1}";
            }

            // Set a print area that spans many rows to ensure pagination
            worksheet.PageSetup.PrintArea = "A1:A500";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use the PaginatedSaveOptions.PageCount property to limit pages
            pdfOptions.PageCount = maxPages;

            // Attach a callback to monitor page saving events
            pdfOptions.PageSavingCallback = new PageCountCallback(maxPages);

            // Save the workbook as PDF (lifecycle create → save)
            string outputPath = "LimitedPagesOutput.pdf";
            workbook.Save(outputPath, pdfOptions);
        }
    }

    // Callback implementation to verify page counts during the save process
    class PageCountCallback : IPageSavingCallback
    {
        private readonly int _maxPages;
        private int _totalPages = -1;

        public PageCountCallback(int maxPages)
        {
            _maxPages = maxPages;
        }

        // Called before each page is saved
        public void PageStartSaving(PageStartSavingArgs args)
        {
            // Capture the total page count on the first page (uses PageSavingArgs.PageCount)
            if (args.PageIndex == 0)
            {
                _totalPages = args.PageCount;
                Console.WriteLine($"Total pages in workbook: {_totalPages}");
                Console.WriteLine($"Maximum pages allowed: {_maxPages}");
            }

            Console.WriteLine($"Saving page {args.PageIndex + 1} of {_totalPages}");
        }

        // Called after each page is saved
        public void PageEndSaving(PageEndSavingArgs args)
        {
            // If the maximum page limit is reached, stop further page processing
            if (args.PageIndex + 1 >= _maxPages)
            {
                args.HasMorePages = false; // Uses PageEndSavingArgs.HasMorePages
                Console.WriteLine($"Reached max page limit ({_maxPages}). No more pages will be saved.");
            }
        }
    }
}
