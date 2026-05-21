using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class SetCustomMarginsToTiff
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample content
        worksheet.Cells["A1"].PutValue("Custom margins demo");

        // Set custom page margins (in inches)
        worksheet.PageSetup.LeftMarginInch = 0.5;
        worksheet.PageSetup.RightMarginInch = 0.5;
        worksheet.PageSetup.TopMarginInch = 0.7;
        worksheet.PageSetup.BottomMarginInch = 0.7;

        // Configure image options for TIFF rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Tiff;
        options.OnePagePerSheet = true; // render the whole sheet on a single page

        // Create a SheetRender with the worksheet and the configured options
        SheetRender renderer = new SheetRender(worksheet, options);

        // Render the worksheet to a TIFF file
        string outputFile = "CustomMarginsOutput.tiff";
        using (FileStream tiffStream = new FileStream(outputFile, FileMode.Create))
        {
            renderer.ToTiff(tiffStream);
        }

        Console.WriteLine($"TIFF file saved to: {outputFile}");
    }
}