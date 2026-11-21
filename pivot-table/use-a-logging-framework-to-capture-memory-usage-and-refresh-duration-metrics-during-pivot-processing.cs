// Title: Log memory consumption and refresh duration of an Aspose.Cells pivot table in C#
// Common Searches: C# code to log memory usage before and after refreshing an Aspose.Cells pivot table | how to measure Aspose.Cells pivot table refresh time with Stopwatch | using GC.GetTotalMemory to monitor Excel pivot processing performance in .NET | track pivot table refresh performance metrics with Aspose.Cells API | record pivot refresh duration and memory delta when saving workbook using Aspose.Cells
// Tags: Aspose.Cells pivot refresh performance logging | Stopwatch timing for pivot table refresh C# | GC memory snapshot Aspose.Cells pivot processing | record pivot refresh metrics .NET | Excel pivot memory delta tracking with Aspose.Cells | save workbook after pivot refresh Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The example loads an Excel workbook with Aspose.Cells, checks for a pivot table, captures memory usage before and after calling RefreshData and CalculateData, measures the refresh duration with Stopwatch, logs the memory delta and elapsed time, and finally saves the workbook while handling possible I/O errors.
class PivotProcessingWithMetrics
{
    static void Main(string[] args)
    {
        // Input and output file paths
        string inputPath = @"C:\Data\InputWorkbook.xlsx";
        string outputPath = @"C:\Data\OutputWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Console.WriteLine($"Workbook loaded from {inputPath}");

            // Assume the first worksheet contains the pivot table
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for pivot tables
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];
            Console.WriteLine($"Pivot table \"{pivotTable.Name}\" found. Starting refresh...");

            // Capture memory usage before refresh
            long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
            Console.WriteLine($"Memory before refresh: {memoryBefore} bytes");

            // Measure refresh duration
            Stopwatch sw = Stopwatch.StartNew();
            pivotTable.RefreshData();
            pivotTable.CalculateData();
            sw.Stop();

            // Capture memory usage after refresh
            long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
            Console.WriteLine($"Memory after refresh: {memoryAfter} bytes");

            // Compute differences
            long memoryDelta = memoryAfter - memoryBefore;
            TimeSpan duration = sw.Elapsed;

            // Log metrics
            Console.WriteLine("Pivot refresh completed.");
            Console.WriteLine($"Refresh duration: {duration.TotalMilliseconds} ms");
            Console.WriteLine($"Memory delta: {memoryDelta} bytes ({memoryDelta / 1024.0:F2} KB)");

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine("Failed to save the workbook:");
                Console.Error.WriteLine(saveEx);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred during pivot processing:");
            Console.Error.WriteLine(ex);
        }
    }
}
