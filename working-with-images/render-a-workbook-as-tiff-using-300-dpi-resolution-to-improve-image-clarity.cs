// Title: Export an Aspose.Cells Workbook to a High‑Resolution Multi‑Page TIFF (300 DPI) in C#
// Description: Creates a workbook, fills sample data, configures ImageOrPrintOptions for TIFF output with 300 DPI horizontal and vertical resolution and LZW compression, then uses WorkbookRender to generate a multi‑page TIFF file.
// Keywords: Aspose.Cells TIFF export C# | 300 DPI TIFF Aspose | WorkbookRender multi‑page TIFF | ImageOrPrintOptions resolution | LZW compression TIFF .NET | Excel to high‑resolution image | Aspose.Cells render options | C# export Excel as TIFF
// Common Searches: how to save Aspose.Cells workbook as 300 DPI TIFF | Aspose.Cells render multi‑page TIFF with LZW | C# set horizontal and vertical resolution for TIFF export | export Excel sheet to high‑resolution TIFF using Aspose | Aspose.Cells TIFF compression options
// Developer Intent: Generate a high‑resolution, multi‑page TIFF image from an Excel workbook using Aspose.Cells in .NET.
// Use Cases: Print‑ready reports that require 300 DPI TIFF for publishing. | Archiving spreadsheets as lossless TIFF images for compliance. | Creating thumbnail previews of worksheets for document portals.
// AI Prompts: Show how to render only the second worksheet to a 300 DPI TIFF. | Give code that saves each worksheet as a separate TIFF file with individual DPI settings. | Explain how to switch the TIFF compression to CCITT Group 4 while keeping 300 DPI resolution.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffRenderDemo
{
    // Creates a workbook, fills sample data, configures ImageOrPrintOptions for TIFF output with 300 DPI horizontal and vertical resolution and LZW compression, then uses WorkbookRender to generate a multi‑page TIFF file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67890);

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Specify TIFF as the output image type
                ImageType = ImageType.Tiff,
                // Set resolution to 300 DPI for higher clarity
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Optional: choose a compression method (LZW is lossless)
                TiffCompression = TiffCompression.CompressionLZW
            };

            // Create a workbook renderer with the configured options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the entire workbook to a multi‑page TIFF file
            string outputPath = "RenderedWorkbook.tiff";
            renderer.ToImage(outputPath);

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine($"Workbook successfully rendered to TIFF at: {outputPath}");
        }
    }
}
