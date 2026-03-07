using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetToImageDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create
            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(456);

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png, // output format
                OnePagePerSheet = true                         // render each sheet as a single page
            };

            // Render the worksheet to an image file using SheetRender
            SheetRender renderer = new SheetRender(sheet, options);
            try
            {
                // Render the first (and only) page to a PNG file
                string outputPath = Path.Combine("output", "WorksheetImage.png");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
                renderer.ToImage(0, outputPath);
                Console.WriteLine($"Worksheet rendered successfully to: {outputPath}");
            }
            finally
            {
                // Release resources used by the renderer
                renderer.Dispose();
            }

            // Optionally save the original workbook for reference
            string workbookPath = Path.Combine("output", "OriginalWorkbook.xlsx");
            workbook.Save(workbookPath);
            Console.WriteLine($"Original workbook saved to: {workbookPath}");
        }
    }
}