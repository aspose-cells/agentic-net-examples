// Title: Measure memory consumption and load time of an Excel workbook with ParsingPivotCachedRecords enabled using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets LoadOptions.ParsingPivotCachedRecords to true, captures memory before and after creating a Workbook, and prints the memory difference. | Provide a snippet that records elapsed milliseconds with Stopwatch and memory delta with GC.GetTotalMemory while loading an .xlsx using Aspose.Cells with pivot cache parsing turned on. | Demonstrate how to force a full garbage collection, establish a baseline memory reading, load a workbook with pivot cache records parsed, and output the consumed memory in KB.
// Common Searches: how to benchmark memory usage of Aspose.Cells workbook load with ParsingPivotCachedRecords | C# measure memory consumption when loading Excel file with pivot cache parsing enabled | Aspose.Cells load options parsingpivotcachedrecords impact on performance | profile workbook loading time and memory in .NET using GC.GetTotalMemory and Stopwatch
// Tags: memory measurement Aspose.Cells LoadOptions | ParsingPivotCachedRecords performance analysis | workbook load time profiling C# | GC.GetTotalMemory Excel load profiling | Stopwatch timing Aspose.Cells workbook

using System;
using System.Diagnostics;
using Aspose.Cells;

// The example loads an .xlsx file with LoadOptions.ParsingPivotCachedRecords set to true, measures memory before and after loading using GC.GetTotalMemory, times the operation with Stopwatch, outputs load time and memory used, and saves the workbook.
class MemoryMeasurementDemo
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string inputPath = "input.xlsx";
        // Optional output path to verify successful load/save
        string outputPath = "output.xlsx";

        // Ensure a clean memory baseline
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Record memory before loading
        long memoryBefore = GC.GetTotalMemory(true);
        Stopwatch timer = Stopwatch.StartNew();

        // Create load options and enable parsing of pivot cached records
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingPivotCachedRecords = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        timer.Stop();
        // Record memory after loading
        long memoryAfter = GC.GetTotalMemory(true);
        long memoryUsed = memoryAfter - memoryBefore;

        Console.WriteLine($"Load time: {timer.ElapsedMilliseconds} ms");
        Console.WriteLine($"Memory consumed during load: {memoryUsed / 1024} KB");

        // Save the workbook (optional, demonstrates normal lifecycle)
        workbook.Save(outputPath);
    }
}
