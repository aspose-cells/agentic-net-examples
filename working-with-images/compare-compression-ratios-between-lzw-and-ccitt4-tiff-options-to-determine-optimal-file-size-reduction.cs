using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // SheetRender, ImageOrPrintOptions, TiffCompression, ColorDepth

class TiffCompressionComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            for (int row = 0; row < 100; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Common image rendering options (default format is TIFF)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // ---------- LZW compression ----------
            options.TiffCompression = TiffCompression.CompressionLZW;
            // LZW works with default color depth (24bpp)
            byte[] lzwBytes;
            using (MemoryStream lzwStream = new MemoryStream())
            {
                SheetRender lzwRenderer = new SheetRender(sheet, options);
                lzwRenderer.ToTiff(lzwStream);          // render to stream
                lzwBytes = lzwStream.ToArray();         // capture byte array
            }

            // ---------- CCITT4 compression ----------
            options.TiffCompression = TiffCompression.CompressionCCITT4;
            // CCITT4 requires 1‑bit per pixel
            options.TiffColorDepth = ColorDepth.Format1bpp;
            byte[] ccitt4Bytes;
            using (MemoryStream ccitt4Stream = new MemoryStream())
            {
                SheetRender ccitt4Renderer = new SheetRender(sheet, options);
                ccitt4Renderer.ToTiff(ccitt4Stream);
                ccitt4Bytes = ccitt4Stream.ToArray();
            }

            // Compare file sizes
            long lzwSize = lzwBytes.Length;
            long ccitt4Size = ccitt4Bytes.Length;

            Console.WriteLine($"LZW TIFF size: {lzwSize} bytes");
            Console.WriteLine($"CCITT4 TIFF size: {ccitt4Size} bytes");

            if (lzwSize < ccitt4Size)
                Console.WriteLine("LZW provides better compression (smaller file).");
            else if (ccitt4Size < lzwSize)
                Console.WriteLine("CCITT4 provides better compression (smaller file).");
            else
                Console.WriteLine("Both compression methods result in the same file size.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}