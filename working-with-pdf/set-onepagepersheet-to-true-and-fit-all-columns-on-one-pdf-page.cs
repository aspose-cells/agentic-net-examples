using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to demonstrate many columns
            for (int col = 0; col < 50; col++)
            {
                sheet.Cells[0, col].PutValue("Header " + (col + 1));
                sheet.Cells[1, col].PutValue("Data " + (col + 1));
            }

            // Optional: set page setup to fit columns on one page
            sheet.PageSetup.FitToPagesWide = 1;   // fit all columns horizontally
            sheet.PageSetup.FitToPagesTall = 0;   // let height adjust automatically

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OnePagePerSheet = true;                // one page per sheet
            pdfOptions.AllColumnsInOnePagePerSheet = true;    // fit all columns on that page

            // Save the workbook as PDF with the specified options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
        }
    }
}