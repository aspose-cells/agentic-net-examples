// Title: Measure .NET Memory Usage During Worksheet‑to‑TIFF Conversion with Aspose.Cells (C#)
// Description: A C# console example that creates a workbook, populates 200 rows, sets TIFF rendering options (LZW compression), logs managed heap size with GC.GetTotalMemory before and after calling SheetRender.ToImage, and prints the memory delta. Useful for assessing the resource impact of Excel‑to‑TIFF export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells TIFF conversion memory | C# GC.GetTotalMemory | measure .NET memory usage | Excel to TIFF performance | worksheet rendering memory profiling | Aspose.Cells image export benchmark
// Common Searches: how to log memory before and after Aspose.Cells TIFF export | measure memory impact of Excel to TIFF conversion .NET | GC.GetTotalMemory usage with Aspose.Cells rendering | benchmark Aspose.Cells TIFF compression memory | track memory consumption during worksheet image export
// Developer Intent: Evaluate the amount of managed memory consumed when converting an Excel worksheet to a TIFF image using Aspose.Cells.
// Use Cases: Profile memory requirements for large workbooks before batch TIFF export. | Detect memory leaks in a long‑running service that repeatedly renders worksheets to images. | Compare memory footprints of different TIFF compression settings (LZW, CCITT, etc.) in Aspose.Cells.
// AI Prompts: Write a reusable C# method that accepts a Worksheet and ImageOrPrintOptions, logs GC.GetTotalMemory before and after rendering, and returns the memory delta. | Show how to extend the example to record memory usage for each sheet in a multi‑sheet workbook and output a summary report. | Explain how GC.GetTotalMemory works and how to interpret its values when measuring Aspose.Cells image export performance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# console example that creates a workbook, populates 200 rows, sets TIFF rendering options (LZW compression), logs managed heap size with GC.GetTotalMemory before and after calling SheetRender.ToImage, and prints the memory delta. Useful for assessing the resource impact of Excel‑to‑TIFF export in Aspose.Cells for .NET.
class TiffMemoryLogger
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            for (int i = 0; i < 200; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Sample row {i}");
            }

            // Set up image rendering options for TIFF
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                TiffCompression = TiffCompression.CompressionLZW
                // ImageFormat is inferred from the output file extension, so it can be omitted
            };

            // Capture memory usage before conversion
            long memoryBefore = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory before TIFF conversion: {memoryBefore} bytes");

            // Prepare output path
            string outputPath = "worksheet_output.tiff";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the worksheet to a TIFF file
            SheetRender renderer = new SheetRender(worksheet, options);
            renderer.ToImage(0, outputPath);

            // Capture memory usage after conversion
            long memoryAfter = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory after TIFF conversion: {memoryAfter} bytes");
            Console.WriteLine($"Memory delta: {memoryAfter - memoryBefore} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
