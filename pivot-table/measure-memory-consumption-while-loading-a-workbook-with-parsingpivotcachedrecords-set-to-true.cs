// Title: Profile memory consumption while loading an XLSX workbook with ParsingPivotCachedRecords enabled in Aspose.Cells for .NET
// Common Searches: profile memory consumption of Aspose.Cells Workbook load with ParsingPivotCachedRecords set to true | measure memory impact of parsing pivot cache records in Aspose.Cells using C# | benchmark Excel file loading time and memory when enabling ParsingPivotCachedRecords in .NET | use GC.GetTotalMemory to monitor Aspose.Cells load performance for pivot tables | determine memory overhead of Aspose.Cells LoadOptions ParsingPivotCachedRecords
// Tags: ParsingPivotCachedRecords memory profiling | Aspose.Cells workbook load performance | Excel workbook memory usage C# | pivot cache parsing benchmark Aspose.Cells | Aspose.Cells memory measurement

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads an XLSX workbook with LoadOptions.ParsingPivotCachedRecords set to true, captures memory before and after loading using GC.GetTotalMemory, measures load duration with Stopwatch, and outputs both elapsed time and memory consumed in kilobytes.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string filePath = "input.xlsx";

        // Measure memory before loading
        long memoryBefore = GC.GetTotalMemory(true);

        // Start timing the load operation
        Stopwatch sw = Stopwatch.StartNew();

        // Configure load options to parse pivot cached records
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.ParsingPivotCachedRecords = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Stop timing
        sw.Stop();

        // Measure memory after loading
        long memoryAfter = GC.GetTotalMemory(true);
        long memoryUsed = memoryAfter - memoryBefore;

        // Output results
        Console.WriteLine($"Load time: {sw.ElapsedMilliseconds} ms");
        Console.WriteLine($"Memory used: {memoryUsed / 1024} KB");
    }
}
