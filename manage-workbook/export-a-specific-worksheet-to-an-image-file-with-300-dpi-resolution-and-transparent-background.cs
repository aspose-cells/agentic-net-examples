using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsExportWorksheet
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create rule

            // Access the worksheet you want to export (e.g., the first worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // OPTIONAL: add some data so the image is not empty
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B2"].PutValue(12345);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Png;          // PNG supports transparency
            options.HorizontalResolution = 300;         // 300 DPI horizontal
            options.VerticalResolution = 300;           // 300 DPI vertical
            options.Transparent = true;                // make background transparent
            options.OnePagePerSheet = true;             // render the whole sheet on one page

            // Create a SheetRender instance for the selected worksheet
            SheetRender sheetRender = new SheetRender(worksheet, options); // constructor rule

            // Define output image path
            string outputPath = Path.Combine(Environment.CurrentDirectory, "WorksheetImage.png");

            // Render the first (and only) page to an image file
            sheetRender.ToImage(0, outputPath); // ToImage(pageIndex, fileName) rule

            Console.WriteLine($"Worksheet exported to image: {outputPath}");
        }
    }
}