using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(150);
        worksheet.Cells["A3"].PutValue("Orange");
        worksheet.Cells["B3"].PutValue(200);

        // Configure image rendering options (PNG format, one page per sheet)
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        // Render the worksheet to a PNG image file using SheetRender
        SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
        string pngFileName = "WorksheetImage.png";
        sheetRender.ToImage(0, pngFileName); // Render first (and only) page

        // Generate an HTML <img> tag that references the saved PNG file
        string htmlImgTag = $"<img src=\"{pngFileName}\" alt=\"Worksheet Image\" />";
        Console.WriteLine(htmlImgTag);
    }
}