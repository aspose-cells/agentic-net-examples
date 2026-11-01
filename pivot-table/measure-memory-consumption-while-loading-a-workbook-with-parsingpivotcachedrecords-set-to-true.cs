// Title: Measure memory usage when loading an Excel workbook with ParsingPivotCachedRecords enabled using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with LoadOptions.ParsingPivotCachedRecords set to true and prints the memory delta using GC.GetTotalMemory. | Create a reusable method that accepts a file path, loads the workbook with pivot cache parsing turned on, and returns the memory consumption in bytes. | Adapt the example to log the before‑and‑after memory values to a file with timestamps instead of writing to the console.
// Common Searches: Aspose.Cells .NET how to benchmark memory consumption of loading a workbook with pivot cache parsing | C# measure memory usage of Workbook constructor when ParsingPivotCachedRecords is true | GC.GetTotalMemory usage example with Aspose.Cells load options | memory impact of ParsingPivotCachedRecords option in Aspose.Cells
// Tags: measure memory usage Aspose.Cells LoadOptions | ParsingPivotCachedRecords memory impact .NET | benchmark workbook loading memory Aspose.Cells | GC.GetTotalMemory Aspose.Cells example | pivot cache parsing performance Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// Demonstrates how to capture memory consumption before and after loading an Excel workbook with the ParsingPivotCachedRecords option enabled, using forced garbage collections and GC.GetTotalMemory to calculate the delta.
class MemoryMeasurement
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Force a full garbage collection and get the initial memory usage (in bytes)
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryBefore = GC.GetTotalMemory(true);

        // Configure load options to parse pivot cached records
        LoadOptions loadOptions = new LoadOptions
        {
            ParsingPivotCachedRecords = true
        };

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Force a garbage collection again and get the memory usage after loading
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryAfter = GC.GetTotalMemory(true);

        // Calculate the memory consumed by the loading operation
        long memoryConsumed = memoryAfter - memoryBefore;

        // Output the result
        Console.WriteLine($"Memory consumed while loading workbook: {memoryConsumed} bytes");
    }
}
