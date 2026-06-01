using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering; // for ImageOrPrintOptions, ColorDepth, TiffCompression

namespace AsposeCellsTiffDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // 2. Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF 24‑bit Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67890);

            // 3. Configure image rendering options for TIFF with 24‑bit color depth
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                 // Specify TIFF output
                TiffColorDepth = ColorDepth.Format24bpp,    // 24‑bit color depth
                TiffCompression = TiffCompression.CompressionLZW // Optional: LZW compression
            };

            // 4. Render the entire workbook to a multi‑page TIFF file (save rule)
            WorkbookRender renderer = new WorkbookRender(workbook, options);
            renderer.ToImage("Workbook_24bit.tiff");

            Console.WriteLine("Workbook successfully rendered to 24‑bit TIFF: Workbook_24bit.tiff");
        }
    }
}