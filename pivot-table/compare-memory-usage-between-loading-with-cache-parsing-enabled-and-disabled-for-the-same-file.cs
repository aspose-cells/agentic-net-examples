// Title: Measure memory usage of Aspose.Cells Workbook loading with default caching vs MemoryPreference setting in C#
// Common Searches: C# how to measure memory consumption when loading an Excel workbook with Aspose.Cells default options | Aspose.Cells load workbook with MemorySetting.MemoryPreference memory usage comparison | compare Aspose.Cells default load options memory impact versus reduced memory setting | benchmark GC.GetTotalMemory for Aspose.Cells workbook loading with and without cache | C# Aspose.Cells memory usage when disabling cache parsing while opening Excel file
// Tags: Aspose.Cells workbook loading memory usage | LoadOptions MemorySetting MemoryPreference | Excel file loading memory optimization | GC.GetTotalMemory benchmark Aspose.Cells | disable cache parsing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads the same Excel file twice using Aspose.Cells for .NET: first with the default LoadOptions (caching enabled) and then with LoadOptions.MemorySetting set to MemoryPreference to lower memory consumption. GC.GetTotalMemory is called before and after each load to calculate the bytes allocated, and the results are printed to the console.
class MemoryComparison
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Ensure the file exists before attempting to load
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // -------------------------------------------------
            // Load with default options (caching enabled by default)
            // -------------------------------------------------
            var defaultLoadOptions = new LoadOptions(); // default behavior includes caching
            long memoryBeforeDefault = GC.GetTotalMemory(true);
            var workbookDefault = new Workbook(filePath, defaultLoadOptions);
            long memoryAfterDefault = GC.GetTotalMemory(true);
            long memoryUsedDefault = memoryAfterDefault - memoryBeforeDefault;

            // -------------------------------------------------
            // Load with reduced memory usage using MemorySetting
            // -------------------------------------------------
            var reducedLoadOptions = new LoadOptions
            {
                // Prefer lower memory consumption (may affect performance)
                MemorySetting = MemorySetting.MemoryPreference
            };
            long memoryBeforeReduced = GC.GetTotalMemory(true);
            var workbookReduced = new Workbook(filePath, reducedLoadOptions);
            long memoryAfterReduced = GC.GetTotalMemory(true);
            long memoryUsedReduced = memoryAfterReduced - memoryBeforeReduced;

            // Output the memory usage comparison
            Console.WriteLine($"Memory used with default loading : {memoryUsedDefault} bytes");
            Console.WriteLine($"Memory used with reduced memory : {memoryUsedReduced} bytes");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
