// Title: C# – Compare LZW vs CCITT4 TIFF Compression with Aspose.Cells
// Description: Learn how to render an Excel workbook to TIFF twice—once with LZW and once with CCITT4 compression—using Aspose.Cells for .NET, capture each stream size, and identify the method that yields the smallest file.
// Keywords: Aspose.Cells TIFF compression | C# LZW vs CCITT4 | TIFF file size comparison | ImageOrPrintOptions compression | optimize TIFF output .NET | Excel to TIFF conversion | high‑resolution TIFF rendering
// Common Searches: Aspose.Cells compare LZW and CCITT4 TIFF size | C# code to measure TIFF compression ratio | which TIFF compression gives smaller files in .NET | render Excel sheet to TIFF with Aspose.Cells | how to get TIFF file size from MemoryStream
// Developer Intent: Find the most space‑efficient TIFF compression (LZW or CCITT4) when converting an Excel workbook to an image with Aspose.Cells.
// Use Cases: Choose the optimal compression for archival TIFF reports generated from Excel data. | Automate batch conversion of worksheets to the smallest possible TIFF files. | Benchmark compression impact on high‑resolution worksheet images before deployment.
// AI Prompts: Create C# code that calculates compression ratios for LZW and CCITT4 TIFF outputs using Aspose.Cells and displays the percentage difference. | Extend the sample to include Deflate and JPEG compression types and generate a summary table of file sizes. | Explain how to configure ImageOrPrintOptions for multi‑page TIFF creation while applying different compression methods per page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering; // for ImageOrPrintOptions and TiffCompression

// Learn how to render an Excel workbook to TIFF twice—once with LZW and once with CCITT4 compression—using Aspose.Cells for .NET, capture each stream size, and identify the method that yields the smallest file.
class TiffCompressionComparison
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Compression Comparison");
        for (int row = 2; row <= 20; row++)
        {
            sheet.Cells[row - 1, 0].PutValue($"Row {row}");
            sheet.Cells[row - 1, 1].PutValue(row * 10);
        }

        // Common image rendering options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            OnePagePerSheet = true // render whole sheet as one page
        };

        // Render with LZW compression
        options.TiffCompression = TiffCompression.CompressionLZW;
        long lzwSize;
        using (MemoryStream lzwStream = new MemoryStream())
        {
            new WorkbookRender(workbook, options).ToImage(lzwStream);
            lzwSize = lzwStream.Length;
        }

        // Render with CCITT4 compression
        options.TiffCompression = TiffCompression.CompressionCCITT4;
        long ccitt4Size;
        using (MemoryStream ccitt4Stream = new MemoryStream())
        {
            new WorkbookRender(workbook, options).ToImage(ccitt4Stream);
            ccitt4Size = ccitt4Stream.Length;
        }

        // Output the file sizes and determine which compression is more effective
        Console.WriteLine($"LZW TIFF size   : {lzwSize} bytes");
        Console.WriteLine($"CCITT4 TIFF size: {ccitt4Size} bytes");

        if (lzwSize < ccitt4Size)
        {
            Console.WriteLine("LZW compression yields a smaller file.");
        }
        else if (ccitt4Size < lzwSize)
        {
            Console.WriteLine("CCITT4 compression yields a smaller file.");
        }
        else
        {
            Console.WriteLine("Both compressions result in the same file size.");
        }
    }
}
