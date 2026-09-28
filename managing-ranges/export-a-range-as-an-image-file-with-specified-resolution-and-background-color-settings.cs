// Title: Export a specific worksheet range to a 300 dpi PNG image using Aspose.Cells for .NET
// AI Prompts: Write C# code that exports cells A1:B3 from a worksheet to a 300 dpi PNG file with Aspose.Cells. | Demonstrate how to set the worksheet print area and configure ImageOrPrintOptions to render a selected range as an image in Aspose.Cells. | Build a reusable C# method that accepts a range address, DPI value, and output path, then saves the range as a PNG using Aspose.Cells.
// Common Searches: how to export a cell range to PNG with Aspose.Cells C# | Aspose.Cells set DPI for image export of worksheet range | C# render selected Excel cells as high resolution PNG using Aspose.Cells | using print area to export specific range as image Aspose.Cells | Aspose.Cells ImageOrPrintOptions export range to PNG example
// Tags: export range to PNG Aspose.Cells | set image DPI Aspose.Cells | configure print area Aspose.Cells | render worksheet range as image C# | high‑resolution Excel image export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Exports the cells A1:B3 of a workbook to a 300 dpi PNG image using Aspose.Cells, configuring the print area and image options before rendering.
class ExportRangeAsImage
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // empty workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for demonstration
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["A3"].PutValue(789);
            sheet.Cells["B3"].PutValue(101112);

            // Define the range to export (e.g., A1:B3)
            string rangeAddress = "A1:B3";

            // Set the print area to the desired range – this tells the renderer which cells to include
            sheet.PageSetup.PrintArea = rangeAddress;

            // Configure image export options.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true
                // BackgroundColor is not available in older versions; the default background will be used.
            };

            // Create a SheetRender object with the worksheet and the image options
            SheetRender renderer = new SheetRender(sheet, imgOptions);

            // Determine output path and ensure the directory exists
            string outputPath = "ExportedRange.png";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the first (and only) page to an image file
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Range exported successfully as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
