using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Memory usage assessment for TIFF conversion");
        // Populate additional cells to simulate a realistic workload
        for (int row = 0; row < 500; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                worksheet.Cells[row, col].PutValue(row * col);
            }
        }

        // Configure rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,
            OnePagePerSheet = true
        };

        // Initialize the sheet renderer (uses the provided ToTiff rule)
        SheetRender renderer = new SheetRender(worksheet, options);

        // Capture memory usage before conversion
        long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
        Console.WriteLine($"Memory before TIFF conversion: {memoryBefore} bytes");

        // Perform the TIFF conversion and write to a file via a stream
        using (MemoryStream tiffStream = new MemoryStream())
        {
            renderer.ToTiff(tiffStream); // Rule: SheetRender.ToTiff(Stream)
            File.WriteAllBytes("output.tiff", tiffStream.ToArray());
        }

        // Capture memory usage after conversion
        long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
        Console.WriteLine($"Memory after TIFF conversion: {memoryAfter} bytes");
        Console.WriteLine($"Memory delta: {memoryAfter - memoryBefore} bytes");
    }
}