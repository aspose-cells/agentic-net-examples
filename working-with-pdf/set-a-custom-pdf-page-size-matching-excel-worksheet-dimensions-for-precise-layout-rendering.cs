using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, just to have content)
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Data");

        // Define the print area that covers the used range
        sheet.PageSetup.PrintArea = "A1:B2";

        // Remove all margins so the page size matches the content exactly
        sheet.PageSetup.LeftMargin = 0;
        sheet.PageSetup.RightMargin = 0;
        sheet.PageSetup.TopMargin = 0;
        sheet.PageSetup.BottomMargin = 0;

        // Render the sheet to obtain the page size in inches
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            OnePagePerSheet = true,
            ImageType = ImageType.Png // required by constructor, not used for PDF
        };
        SheetRender sheetRender = new SheetRender(sheet, renderOptions);
        float[] pageSizeInInches = sheetRender.GetPageSizeInch(0); // [0]=width, [1]=height

        // Set a custom paper size that matches the calculated dimensions
        sheet.PageSetup.CustomPaperSize(pageSizeInInches[0], pageSizeInInches[1]);

        // Save the workbook as PDF; the custom page size will be applied
        workbook.Save("CustomSizeWorksheet.pdf", SaveFormat.Pdf);
    }
}

// Author: Example demonstrating how to set a custom PDF page size that matches the worksheet dimensions.