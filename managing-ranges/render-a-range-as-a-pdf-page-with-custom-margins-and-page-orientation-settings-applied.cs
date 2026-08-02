using System;
using Aspose.Cells;

class RenderRangeToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data in the worksheet
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["C1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["C2"].PutValue(0.5);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["C3"].PutValue(0.3);
        worksheet.Cells["A4"].PutValue("Cherry");
        worksheet.Cells["B4"].PutValue(15);
        worksheet.Cells["C4"].PutValue(0.8);

        // Define the range to be rendered as a PDF page
        // Setting the print area ensures only this range is saved
        worksheet.PageSetup.PrintArea = "A1:C4";

        // Apply custom margins (values are in inches)
        worksheet.PageSetup.TopMarginInch = 0.5;      // top margin
        worksheet.PageSetup.BottomMarginInch = 0.5;   // bottom margin
        worksheet.PageSetup.LeftMarginInch = 0.75;    // left margin
        worksheet.PageSetup.RightMarginInch = 0.75;   // right margin

        // Set the page orientation (Portrait or Landscape)
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Create PDF save options; ensure the range fits on a single page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true   // forces the print area onto one PDF page
        };

        // Save the workbook as a PDF file with the custom settings applied
        workbook.Save("RangeRendered.pdf", pdfOptions);
    }
}