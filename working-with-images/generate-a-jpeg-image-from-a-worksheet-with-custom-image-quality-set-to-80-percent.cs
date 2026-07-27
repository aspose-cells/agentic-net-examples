using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageQualityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Jpeg;   // Output format JPEG
            options.Quality = 80;                // Set JPEG quality to 80%

            // Create a SheetRender instance with the worksheet and options
            SheetRender renderer = new SheetRender(sheet, options);

            // Render the first page of the worksheet to a JPEG file
            renderer.ToImage(0, "Worksheet_Output_Quality80.jpg");

            Console.WriteLine("Worksheet rendered to JPEG with quality 80%.");
        }
    }
}