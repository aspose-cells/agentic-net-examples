using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Sample data for 24-bit TIFF conversion");
        ws.Cells["B2"].PutValue(12345);

        // Configure image options for TIFF with 24‑bit color depth
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            TiffColorDepth = ColorDepth.Format24bpp,
            TiffCompression = TiffCompression.CompressionLZW // optional, default is LZW
        };

        // Create a workbook renderer using the configured options
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Save the entire workbook as a multi‑page TIFF file (lifecycle save)
        renderer.ToImage("output_24bit.tiff");

        // Release resources
        renderer.Dispose();

        Console.WriteLine("Workbook successfully saved as 24‑bit TIFF.");
    }
}