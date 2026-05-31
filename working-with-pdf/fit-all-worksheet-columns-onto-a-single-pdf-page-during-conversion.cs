using System;
using Aspose.Cells;

class FitAllColumnsToOnePdfPage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data across many columns to demonstrate fitting
        for (int col = 0; col < 50; col++)
        {
            worksheet.Cells[0, col].PutValue("Column " + (col + 1));
            worksheet.Cells[1, col].PutValue("Sample data " + (col + 1));
        }

        // Optional: set page setup to fit all columns wide (height will adjust automatically)
        worksheet.PageSetup.FitToPagesWide = 1;   // fit to 1 page wide
        worksheet.PageSetup.FitToPagesTall = 0;   // let height scale automatically

        // Configure PDF save options to force all columns onto a single page per sheet
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.OnePagePerSheet = true;               // one page per sheet
        pdfSaveOptions.AllColumnsInOnePagePerSheet = true;   // all columns on that page

        // Save the workbook as PDF using the configured options
        workbook.Save("AllColumnsOnePage.pdf", pdfSaveOptions);
    }
}