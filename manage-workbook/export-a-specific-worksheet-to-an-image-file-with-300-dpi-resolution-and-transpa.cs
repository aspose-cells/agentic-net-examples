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
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Select the worksheet to export (by index or name)
            Worksheet worksheet = workbook.Worksheets[0]; // change index as needed

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // PNG supports transparency
                HorizontalResolution = 300,         // 300 DPI horizontal
                VerticalResolution = 300,           // 300 DPI vertical
                Transparent = true,                 // make background transparent
                OnePagePerSheet = true              // render the whole sheet on one page
            };

            // Create a SheetRender instance for the selected worksheet
            SheetRender sheetRender = new SheetRender(worksheet, options);

            // Output image file path
            string outputImagePath = "worksheet_export.png";

            // Render the first (and only) page of the sheet to the image file
            sheetRender.ToImage(0, outputImagePath);

            // Clean up resources
            sheetRender.Dispose();

            Console.WriteLine($"Worksheet exported to image: {outputImagePath}");
        }
    }
}