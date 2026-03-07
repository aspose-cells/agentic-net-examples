using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SampleImageFromWorksheet
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(120);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = Aspose.Cells.Drawing.ImageType.Png;
            options.OnePagePerSheet = true; // Render the whole sheet on a single page

            // Create a SheetRender instance for the worksheet
            SheetRender renderer = new SheetRender(sheet, options);

            // Render the first page of the sheet to a PNG file
            string outputPath = "WorksheetImage.png";
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet rendered to image: {outputPath}");
        }
    }
}