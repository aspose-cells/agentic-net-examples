using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class RenderRangeToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item 2");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("Item 3");
        worksheet.Cells["B4"].PutValue(300);

        // Define the range that should be rendered
        string rangeAddress = "A1:B4";

        // Set the print area to the defined range
        worksheet.PageSetup.PrintArea = rangeAddress;

        // Apply custom margins (values are in inches)
        worksheet.PageSetup.TopMarginInch = 0.5;      // top margin
        worksheet.PageSetup.BottomMarginInch = 0.5;   // bottom margin
        worksheet.PageSetup.LeftMarginInch = 0.75;    // left margin
        worksheet.PageSetup.RightMarginInch = 0.75;   // right margin

        // Set the desired page orientation
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Create PDF save options (inherits from PaginatedSaveOptions)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF; only the specified print area will be rendered
        workbook.Save("RenderedRange.pdf", pdfOptions);
    }
}