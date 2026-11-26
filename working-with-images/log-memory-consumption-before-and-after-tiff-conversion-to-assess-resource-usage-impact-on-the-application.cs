// Title: Log memory consumption before and after converting an Excel workbook to TIFF using Aspose.Cells for .NET
// Common Searches: C# measure memory before and after Aspose.Cells workbook to TIFF conversion | Aspose.Cells log memory usage during TIFF rendering in .NET | How to profile memory consumption of Excel to TIFF conversion with Aspose.Cells | Get total memory usage in C# before and after SheetRender ToTiff | Track GC memory delta when converting Excel to TIFF using Aspose.Cells
// Tags: Aspose.Cells memory profiling | SheetRender ToTiff memory usage | ImageOrPrintOptions TIFF conversion performance | GC.GetTotalMemory Aspose.Cells | Excel to TIFF resource monitoring

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // Needed for ImageType enum

namespace AsposeCellsTiffMemoryLogger
{
    // The example loads an Excel workbook (creating a sample file if none exists), forces garbage collection, records total memory, renders the first worksheet to a TIFF file with LZW compression and 300 dpi resolution using ImageOrPrintOptions, forces garbage collection again, records memory after rendering, and prints the before/after memory usage, the delta, and the generated TIFF file size.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.tiff";

            try
            {
                // Ensure the input workbook exists; if not, create a simple one
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                    try
                    {
                        CreateSampleWorkbook(inputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Failed to create sample workbook:");
                        Console.WriteLine(ex.Message);
                        return;
                    }
                }

                // Load the workbook
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Force a full garbage collection before measuring
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
                    Console.WriteLine($"Memory usage before TIFF conversion: {FormatBytes(memoryBefore)}");

                    // Configure TIFF rendering options
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Tiff,
                        TiffCompression = TiffCompression.CompressionLZW,
                        TiffColorDepth = ColorDepth.Format24bpp,
                        HorizontalResolution = 300,
                        VerticalResolution = 300
                    };

                    // Render the first worksheet to TIFF
                    try
                    {
                        SheetRender renderer = new SheetRender(workbook.Worksheets[0], options);
                        renderer.ToTiff(outputPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error during TIFF rendering:");
                        Console.WriteLine(ex.Message);
                        throw; // Re‑throw to be caught by outer handler
                    }

                    // Force another full garbage collection after rendering
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
                    Console.WriteLine($"Memory usage after TIFF conversion:  {FormatBytes(memoryAfter)}");

                    long diff = memoryAfter - memoryBefore;
                    Console.WriteLine($"Memory delta (after - before): {FormatBytes(diff)}");
                }

                // Verify that the TIFF file was created
                if (File.Exists(outputPath))
                {
                    FileInfo fi = new FileInfo(outputPath);
                    Console.WriteLine($"TIFF file generated successfully: \"{outputPath}\" ({FormatBytes(fi.Length)}).");
                }
                else
                {
                    Console.WriteLine("Error: TIFF file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue("Sample Data");
            for (int i = 2; i <= 100; i++)
            {
                ws.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                ws.Cells[$"B{i}"].PutValue(DateTime.Now.AddDays(i - 2).ToString("yyyy-MM-dd"));
                ws.Cells[$"C{i}"].PutValue(i * 10);
            }

            // Apply a simple style to make the sheet a bit richer
            Style style = wb.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10;
            style.Font.IsBold = true;
            ws.Cells.CreateRange("A1:C1").ApplyStyle(style, new StyleFlag { FontBold = true, FontSize = true, FontName = true });

            wb.Save(path, SaveFormat.Xlsx);
            Console.WriteLine($"Sample workbook created at \"{path}\".");
        }

        private static string FormatBytes(long bytes)
        {
            const double ONE_KB = 1024.0;
            const double ONE_MB = ONE_KB * 1024.0;
            const double ONE_GB = ONE_MB * 1024.0;

            if (bytes >= ONE_GB)
                return $"{bytes / ONE_GB:0.##} GB";
            if (bytes >= ONE_MB)
                return $"{bytes / ONE_MB:0.##} MB";
            if (bytes >= ONE_KB)
                return $"{bytes / ONE_KB:0.##} KB";
            return $"{bytes} B";
        }
    }
}
