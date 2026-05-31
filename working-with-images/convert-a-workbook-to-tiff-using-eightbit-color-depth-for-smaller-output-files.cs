using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering; // for ImageOrPrintOptions, ColorDepth, TiffCompression

namespace AsposeCellsTiffConversion
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure image rendering options for TIFF output
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,               // Set output format to TIFF
                TiffColorDepth = ColorDepth.Format8bpp,   // Use 8‑bit color depth for smaller file size
                TiffCompression = TiffCompression.CompressionLZW // Optional: apply LZW compression
            };

            // Create a renderer for the whole workbook with the specified options
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render the workbook to a multi‑page TIFF file using the 8‑bit color depth setting
            renderer.ToImage("output_8bpp.tiff");

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine("Workbook successfully converted to 8‑bit TIFF: output_8bpp.tiff");
        }
    }
}