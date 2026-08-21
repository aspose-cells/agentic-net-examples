// Title: Convert an Excel workbook to a high‑quality 24‑bit multi‑page TIFF with LZW compression in C#
// Description: This example creates a Workbook, adds sample data, sets ImageOrPrintOptions to TIFF with 24‑bit (Format24bpp) color depth and LZW lossless compression, and uses WorkbookRender to save the whole workbook as a multi‑page TIFF file (Workbook24Bit.tiff).
// Keywords: Aspose.Cells C# TIFF export | 24‑bit TIFF Format24bpp | LZW compression Aspose | multi‑page TIFF rendering | WorkbookRender ImageOrPrintOptions | Excel to image .NET | high‑resolution spreadsheet image
// Common Searches: C# Aspose.Cells export workbook to 24‑bit TIFF | How to create multi‑page TIFF from Excel using Aspose | Set TIFF color depth to 24‑bit in Aspose.Cells .NET | LZW compressed TIFF from Excel workbook | Render Excel sheet as high‑quality TIFF image
// Developer Intent: The developer wants to generate a lossless, 24‑bit TIFF image of an entire workbook for archival or documentation purposes.
// Use Cases: Archiving financial statements as printable, lossless TIFF files. | Embedding high‑resolution spreadsheet snapshots in technical manuals. | Creating multi‑page image bundles for legal or compliance records.
// AI Prompts: Show how to set the DPI to 300 while keeping 24‑bit color depth for the TIFF output. | Provide code to write the 24‑bit TIFF to a MemoryStream instead of a file. | Explain how to render only a selected worksheet to a 24‑bit TIFF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffDemo
{
    // This example creates a Workbook, adds sample data, sets ImageOrPrintOptions to TIFF with 24‑bit (Format24bpp) color depth and LZW lossless compression, and uses WorkbookRender to save the whole workbook as a multi‑page TIFF file (Workbook24Bit.tiff).
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF 24‑bit Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // Configure image rendering options for TIFF with 24‑bit color depth
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                 // Ensure TIFF output
                TiffColorDepth = ColorDepth.Format24bpp,    // 24‑bit color depth for high quality
                TiffCompression = TiffCompression.CompressionLZW // Optional: lossless compression
            };

            // Render the entire workbook to a multi‑page TIFF file (lifecycle: save)
            WorkbookRender renderer = new WorkbookRender(workbook, options);
            renderer.ToImage("Workbook24Bit.tiff");

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine("Workbook successfully rendered to 24‑bit TIFF: Workbook24Bit.tiff");
        }
    }
}
