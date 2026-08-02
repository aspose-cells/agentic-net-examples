using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates an empty workbook
            // Example: load an existing file
            // Workbook workbook = new Workbook("input.xlsx");

            // Add some sample data to demonstrate the conversion
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for 8‑bit TIFF conversion");
            sheet.Cells["B2"].PutValue(12345);

            // Configure image rendering options for TIFF with 8‑bit color depth
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,          // Specify TIFF output
                TiffColorDepth = ColorDepth.Format8bpp, // Set 8‑bit per pixel
                TiffCompression = TiffCompression.CompressionLZW // Optional: use LZW compression
            };

            // Render the entire workbook to a single TIFF file using the specified options
            WorkbookRender renderer = new WorkbookRender(workbook, options);
            renderer.ToImage("output_8bpp.tiff");

            Console.WriteLine("Workbook successfully converted to 8‑bit TIFF: output_8bpp.tiff");
        }
    }
}