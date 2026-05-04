using System;
using Aspose.Cells;

namespace AsposeCellsColumnFitPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data (enough columns to require scaling)
            for (int col = 0; col < 30; col++)
            {
                // Header
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                // Sample row data
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Adjust page setup to fit all columns on a single page (width = 1, height = auto)
            sheet.PageSetup.FitToPagesWide = 1;   // one page wide
            sheet.PageSetup.FitToPagesTall = 0;   // let height adjust automatically

            // Create PDF save options and enable one-page-per-sheet behavior
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true;                // all content on one page
            pdfOptions.AllColumnsInOnePagePerSheet = true;    // force all columns onto that page

            // Save the workbook as PDF; the entire sheet will be scaled to fit one page
            workbook.Save("SheetFitOnePage.pdf", pdfOptions);
        }
    }
}