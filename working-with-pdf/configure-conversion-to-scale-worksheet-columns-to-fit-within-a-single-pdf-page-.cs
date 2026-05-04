using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data (optional, just for demonstration)
        for (int col = 0; col < 30; col++)
        {
            worksheet.Cells[0, col].PutValue("Header " + (col + 1));
            worksheet.Cells[1, col].PutValue("Data " + (col + 1));
        }

        // Configure page setup so that all columns are fitted to a single page width.
        // Setting FitToPagesWide = 1 forces the sheet to be scaled to one page wide.
        // Setting FitToPagesTall = 0 lets the height adjust automatically.
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 0;

        // Create PDF save options.
        // OnePagePerSheet = true ensures each sheet is rendered on a single page.
        // AllColumnsInOnePagePerSheet = true forces all column content to fit on that page.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as a PDF with the configured scaling.
        workbook.Save("ScaledColumns.pdf", pdfOptions);
    }
}