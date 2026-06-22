using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class TiffMemoryLogger
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int i = 0; i < 100; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i}");
            worksheet.Cells[i, 1].PutValue(i);
        }

        // Configure TIFF rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            OnePagePerSheet = true,
            TiffCompression = TiffCompression.CompressionLZW,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // Initialize the sheet renderer with the worksheet and options
        SheetRender renderer = new SheetRender(worksheet, options);

        // Log memory usage before conversion
        long memoryBefore = GC.GetTotalMemory(true);
        Console.WriteLine($"Memory before TIFF conversion: {memoryBefore} bytes");

        // Render the worksheet to a TIFF image using a memory stream
        using (MemoryStream tiffStream = new MemoryStream())
        {
            renderer.ToTiff(tiffStream);
            Console.WriteLine($"Generated TIFF size in memory: {tiffStream.Length} bytes");
        }

        // Log memory usage after conversion
        long memoryAfter = GC.GetTotalMemory(true);
        Console.WriteLine($"Memory after TIFF conversion: {memoryAfter} bytes");
        Console.WriteLine($"Memory delta: {memoryAfter - memoryBefore} bytes");
    }
}