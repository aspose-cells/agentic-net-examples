using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ExportSheetsToJpeg
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Configure image rendering options
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,          // Export as JPEG
                    HorizontalResolution = 300,          // 300 DPI horizontal
                    VerticalResolution = 300,            // 300 DPI vertical
                    Quality = 100,                       // Highest JPEG quality
                    OnePagePerSheet = true               // One image per sheet
                };

                // Create a SheetRender for the current worksheet (uses the provided create rule)
                SheetRender renderer = new SheetRender(sheet, options);

                // Generate the JPEG file name (e.g., Sheet_1.jpg, Sheet_2.jpg, ...)
                string outputFileName = $"Sheet_{sheetIndex + 1}.jpg";

                // Render the first (and only) page of the sheet to a JPEG file (uses the provided save rule)
                renderer.ToImage(0, outputFileName);

                Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{outputFileName}' at 300 DPI.");
            }
        }
    }
}