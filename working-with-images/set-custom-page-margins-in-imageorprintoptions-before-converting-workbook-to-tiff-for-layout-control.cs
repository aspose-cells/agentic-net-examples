using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Demo: Custom page margins for TIFF rendering");

        // Set custom page margins (values are in inches)
        worksheet.PageSetup.TopMarginInch = 0.5;      // 0.5 inch top margin
        worksheet.PageSetup.BottomMarginInch = 0.5;   // 0.5 inch bottom margin
        worksheet.PageSetup.LeftMarginInch = 0.75;    // 0.75 inch left margin
        worksheet.PageSetup.RightMarginInch = 0.75;   // 0.75 inch right margin

        // Configure image rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff;          // Set output format to TIFF
        options.OnePagePerSheet = true;              // Render the whole sheet on a single page

        // Create a SheetRender object with the worksheet and options
        SheetRender renderer = new SheetRender(worksheet, options);

        // Render the worksheet to a multi‑page TIFF file
        string outputFile = "CustomMarginsOutput.tiff";
        renderer.ToTiff(outputFile);

        Console.WriteLine($"TIFF file created with custom margins: {outputFile}");
    }
}