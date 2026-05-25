using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfPageCheck
{
    // Callback to capture total page count during PDF generation
    public class PageCountCallback : IPageSavingCallback
    {
        // Holds the total number of pages in the workbook (set once)
        public static int TotalPages { get; private set; } = 0;

        public void PageStartSaving(PageStartSavingArgs args)
        {
            // PageCount is the total pages that will be rendered
            // Capture it the first time the callback is invoked
            if (TotalPages == 0)
            {
                TotalPages = args.PageCount;
            }

            // Optionally, you could stop after a certain page index here
            // but we rely on PdfSaveOptions.PageCount to enforce the limit
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // No additional processing needed after each page
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define the maximum number of pages we want in the PDF
            int maxPages = 5;

            // Create a new workbook and populate it with enough data to span multiple pages
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill 500 rows to ensure pagination
            for (int i = 0; i < 500; i++)
            {
                sheet.Cells[i, 0].Value = $"Row {i + 1}";
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Limit the number of pages that will be saved
                PageCount = maxPages,

                // Attach the callback to capture total page count
                PageSavingCallback = new PageCountCallback()
            };

            // Save the workbook as PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);

            // After saving, retrieve the total page count discovered during rendering
            int totalPages = PageCountCallback.TotalPages;

            // Determine how many pages were actually written (cannot exceed maxPages)
            int pagesSaved = Math.Min(totalPages, maxPages);

            // Output verification results
            Console.WriteLine($"Total pages in workbook (before limit): {totalPages}");
            Console.WriteLine($"Maximum pages allowed: {maxPages}");
            Console.WriteLine($"Pages actually saved to PDF: {pagesSaved}");
        }
    }
}