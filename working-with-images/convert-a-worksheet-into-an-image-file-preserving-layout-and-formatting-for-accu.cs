using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorksheetToImage
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apples");
        worksheet.Cells["B2"].PutValue(150);
        worksheet.Cells["A3"].PutValue("Bananas");
        worksheet.Cells["B3"].PutValue(200);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png, // Output format
            OnePagePerSheet = true                         // Render the whole sheet on one page
        };

        // Create a SheetRender instance for the worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string imagePath = Path.Combine(outputDir, "worksheet.png");

        // Render the first (and only) page of the sheet to an image file
        sheetRender.ToImage(0, imagePath);

        Console.WriteLine($"Worksheet successfully rendered to image: {imagePath}");

        // Release resources used by SheetRender
        sheetRender.Dispose();
    }
}