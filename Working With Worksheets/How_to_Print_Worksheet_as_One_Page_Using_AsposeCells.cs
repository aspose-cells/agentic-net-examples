using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PrintWorksheetOnePage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 100; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure page setup so that all rows fit on a single page vertically.
        // Setting FitToPagesWide = 0 lets the width adjust automatically.
        worksheet.PageSetup.FitToPagesTall = 1;
        worksheet.PageSetup.FitToPagesWide = 0;

        // Create ImageOrPrintOptions and enable OnePagePerSheet.
        // This forces the entire sheet to be rendered on one page.
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };

        // Create SheetRender with the worksheet and the configured options.
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first (and only) page to an image file.
        // Because OnePagePerSheet = true, PageCount will be 1.
        sheetRender.ToImage(0, "WorksheetOnePage.png");

        // If you need to send the sheet directly to a printer, uncomment the line below
        // and replace the printer name with an installed printer on the system.
        // sheetRender.ToPrinter("Microsoft Print to PDF");

        // Release resources
        sheetRender.Dispose();
    }
}