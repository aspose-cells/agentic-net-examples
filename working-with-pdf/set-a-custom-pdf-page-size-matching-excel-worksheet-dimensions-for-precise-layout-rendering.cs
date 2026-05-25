using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class CustomPdfPageSize
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Configure page setup to fit all content on a single page
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.FitToPagesWide = 1;   // fit columns to one page width
        pageSetup.FitToPagesTall = 1;   // fit rows to one page height
        pageSetup.PaperSize = PaperSizeType.Custom; // enable custom size

        // Render the sheet to obtain the calculated page size in inches
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
        SheetRender sheetRender = new SheetRender(sheet, renderOptions);
        float[] pageSizeInches = sheetRender.GetPageSizeInch(0); // [0]=width, [1]=height

        // Apply the custom paper size matching the rendered dimensions
        double widthInches = pageSizeInches[0];
        double heightInches = pageSizeInches[1];
        pageSetup.CustomPaperSize(widthInches, heightInches);

        // Save the workbook as PDF; the custom page size will be used
        workbook.Save("CustomSize.pdf", SaveFormat.Pdf);

        // Clean up
        sheetRender.Dispose();
    }
}