using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – sets OnePagePerSheet and limits PDF to 10 pages
    class OnePagePerSheetWithPageLimit
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 200; i++)
            {
                sheet.Cells[i, 0].Value = $"Row {i + 1}";
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,                     // All content of a sheet fits on one page
                PageSavingCallback = new MaxPageCallback(10) // Limit output to a maximum of 10 pages
            };

            // Save the workbook as PDF
            workbook.Save("OnePagePerSheet_Limited.pdf", pdfOptions);
        }
    }

    // Callback that stops saving after a specified number of pages
    internal class MaxPageCallback : IPageSavingCallback
    {
        private readonly int _maxPages;
        private int _currentPage;

        public MaxPageCallback(int maxPages)
        {
            _maxPages = maxPages;
            _currentPage = 0;
        }

        public void PageStartSaving(PageStartSavingArgs args)
        {
            // No action needed before a page is saved
        }

        public void PageEndSaving(PageEndSavingArgs args)
        {
            // Increment page counter
            _currentPage++;

            // If the maximum page count is reached, prevent further pages
            if (_currentPage >= _maxPages)
            {
                args.HasMorePages = false;
            }
        }
    }
}