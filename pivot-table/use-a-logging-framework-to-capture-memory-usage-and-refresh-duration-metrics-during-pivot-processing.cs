// Title: Capture memory usage and execution time while refreshing an Aspose.Cells pivot table in C#
// Common Searches: how to measure pivot table refresh time with Aspose.Cells in C# | record memory consumption during Aspose.Cells pivot refresh using .NET | log execution duration and memory delta for Excel pivot processing with Aspose.Cells | benchmark Aspose.Cells pivot table refresh performance in a C# application | using GC.GetTotalMemory and Stopwatch to monitor Aspose.Cells pivot updates
// Tags: Aspose.Cells pivot refresh performance monitoring | C# memory measurement with GC.GetTotalMemory for Excel pivot | Stopwatch timing of Aspose.Cells pivot table update | log pivot processing metrics .NET | Excel pivot table refresh memory usage Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies the input file, accesses the first worksheet's pivot table, captures memory usage before and after refreshing the pivot data with GC.GetTotalMemory, measures the refresh duration using Stopwatch, logs both the elapsed milliseconds and memory delta to the console, saves the updated workbook, and includes exception handling.
class PivotProcessingMetrics
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the pivot table
            var worksheet = workbook.Worksheets[0];
            var pivotTables = worksheet.PivotTables;

            if (pivotTables.Count == 0)
            {
                Console.WriteLine("Warning: No pivot tables found in the worksheet.");
                return;
            }

            // Use the first pivot table for demonstration
            var pivotTable = pivotTables[0];

            // Capture memory usage before refresh
            long memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
            var stopwatch = Stopwatch.StartNew();

            // Refresh pivot data and recalculate
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Stop timing and capture memory after refresh
            stopwatch.Stop();
            long memoryAfter = GC.GetTotalMemory(forceFullCollection: true);
            long memoryUsed = memoryAfter - memoryBefore;

            // Log duration and memory metrics
            Console.WriteLine($"Pivot refresh duration: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"Memory used during refresh: {memoryUsed} bytes");

            // Save the workbook with refreshed pivot
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred during pivot processing: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
