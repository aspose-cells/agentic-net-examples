using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class FitColumnsOnePage
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data across many columns
        for (int col = 0; col < 30; col++)
        {
            // Header row
            worksheet.Cells[0, col].PutValue($"Header {col + 1}");

            // Data rows
            for (int row = 1; row <= 20; row++)
            {
                worksheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Configure page setup to fit all columns on a single page width
        // Set FitToPagesWide to 1 (one page wide) and FitToPagesTall to 0 (height adjusts automatically)
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 0;

        // Optional: define a print area covering the used range
        // worksheet.PageSetup.PrintArea = "A1:" + CellsHelper.CellIndexToName(worksheet.Cells.MaxDataColumn, worksheet.Cells.MaxDataRow);

        // Render the worksheet to an image respecting the page setup
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
        renderOptions.OnePagePerSheet = true; // ensures the whole sheet is rendered as one page
        SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

        // Save the rendered page as PNG
        sheetRender.ToImage(0, "FitColumnsOnePage.png");

        // Save the workbook (optional, for verification)
        workbook.Save("FitColumnsOnePage.xlsx");
    }
}