using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class CustomPaperSizeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data to illustrate the page content
        worksheet.Cells["A1"].PutValue("Custom Paper Size Demo");
        worksheet.Cells["A2"].PutValue("This page uses a 4\" x 6\" paper.");

        // Set a custom paper size (width: 4 inches, height: 6 inches)
        worksheet.PageSetup.CustomPaperSize(4.0, 6.0);
        // Ensure the page setup uses the custom size
        worksheet.PageSetup.PaperSize = PaperSizeType.Custom;

        // Export the workbook to PDF – PDF respects the page setup settings
        workbook.Save("CustomPaperSize.pdf", SaveFormat.Pdf);

        // Additionally render the first page to a PNG image
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        WorkbookRender render = new WorkbookRender(workbook, options);
        render.ToImage(0, "CustomPaperSize.png");
        render.Dispose();
    }
}