// Title: Set custom worksheet page margins with ImageOrPrintOptions and export Excel to multi-page TIFF using Aspose.Cells for .NET
// Common Searches: how to set page margins in Aspose.Cells before converting Excel to TIFF in C# | Aspose.Cells ImageOrPrintOptions margin configuration for TIFF output | C# render Excel workbook to multi-page TIFF with custom margins using Aspose.Cells | adjust top, bottom, left, right margins when exporting Excel to TIFF with Aspose.Cells | configure LZW compression and 300 dpi while setting margins for TIFF conversion in Aspose.Cells
// Tags: ImageOrPrintOptions margin configuration | WorkbookRender TIFF export | PageSetup custom margins | TIFF LZW compression Aspose.Cells | multi-page TIFF rendering C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffMarginDemo
{
    // The example loads or creates an Excel workbook, applies 0.5‑inch top/bottom and 0.75‑inch left/right margins to the first worksheet via PageSetup, configures ImageOrPrintOptions for TIFF output (LZW compression, 300 dpi, one page per sheet), and uses WorkbookRender to generate a single multi‑page TIFF file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input Excel file. If the file does not exist, a sample workbook will be created.
            string inputPath = "input.xlsx";
            string outputPath = "output.tiff";

            Workbook workbook = null;

            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                // Create a sample workbook with some data
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["C1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0.5);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(0.3);
                sheet.Cells["A4"].PutValue("Total");
                sheet.Cells["B4"].Formula = "=SUM(B2:B3)";
                sheet.Cells["C4"].Formula = "=SUMPRODUCT(B2:B3,C2:C3)";

                // Save the sample workbook for reference
                workbook.Save(inputPath);
                Console.WriteLine($"Sample workbook created and saved as \"{inputPath}\".");
            }

            // ------------------------------------------------------------
            // Set custom page margins (in inches) on the first worksheet.
            // ------------------------------------------------------------
            Worksheet firstSheet = workbook.Worksheets[0];
            PageSetup pageSetup = firstSheet.PageSetup;

            // Margins are specified in inches. Adjust as needed.
            pageSetup.TopMargin = 0.5;      // 0.5 inch from top
            pageSetup.BottomMargin = 0.5;   // 0.5 inch from bottom
            pageSetup.LeftMargin = 0.75;    // 0.75 inch from left
            pageSetup.RightMargin = 0.75;   // 0.75 inch from right

            Console.WriteLine("Custom page margins applied:");
            Console.WriteLine($"Top: {pageSetup.TopMargin} in, Bottom: {pageSetup.BottomMargin} in, Left: {pageSetup.LeftMargin} in, Right: {pageSetup.RightMargin} in");

            // ------------------------------------------------------------
            // Configure ImageOrPrintOptions for TIFF conversion.
            // ------------------------------------------------------------
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Tiff,
                TiffCompression = TiffCompression.CompressionLZW,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Ensure each sheet is rendered on a single page (optional)
                OnePagePerSheet = true
            };

            // ------------------------------------------------------------
            // Render the entire workbook to a single TIFF file.
            // ------------------------------------------------------------
            try
            {
                WorkbookRender renderer = new WorkbookRender(workbook, options);
                Console.WriteLine($"Total pages to render: {renderer.PageCount}");

                // Render all pages into one multi-page TIFF file.
                renderer.ToImage(outputPath);
                Console.WriteLine($"Workbook successfully rendered to TIFF: \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during TIFF rendering: {ex.Message}");
            }
        }
    }
}
