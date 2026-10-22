// Title: How to compare memory usage of loading an XLSX workbook with cache parsing enabled versus disabled using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads the same .xlsx file twice with Aspose.Cells—first using default LoadOptions and then with LoadOptions.MemorySetting set to MemoryPreference—while measuring memory before and after each load and printing the results. | Show a snippet that forces garbage collection, captures GC.GetTotalMemory, and calculates the memory delta for each workbook load to benchmark cache parsing impact. | Demonstrate how to disable internal caching in Aspose.Cells load options and output a side‑by‑side comparison of the memory footprint against the default loading behavior.
// Common Searches: Aspose.Cells compare memory consumption cache parsing on and off | measure memory impact of disabling cache parsing when loading XLSX in C# | benchmark Aspose.Cells workbook load memory usage with MemorySetting.MemoryPreference | C# how to reduce memory usage of Aspose.Cells by turning off internal cache | difference in RAM usage between default load and memory‑optimized load in Aspose.Cells
// Tags: Aspose.Cells load options memory setting | disable cache parsing Aspose.Cells | benchmark workbook load memory .NET | memory optimization XLSX loading Aspose.Cells | compare cache parsing memory usage

using System;
using System.IO;
using Aspose.Cells;
using System.Diagnostics;

// // Loads the same XLSX file twice using Aspose.Cells: first with default LoadOptions (cache parsing enabled) and then with LoadOptions.MemorySetting = MemoryPreference (cache disabled). The program forces garbage collection before each load, captures GC.GetTotalMemory, computes the memory delta, and prints the memory used for each scenario in kilobytes.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // -------------------------------------------------
            // Load with default memory settings (cache parsing enabled)
            // -------------------------------------------------
            var loadOptionsEnabled = new LoadOptions(LoadFormat.Xlsx);
            // No special settings – default behavior uses internal caching

            // Clean up before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeEnabled = GC.GetTotalMemory(true);

            // Load the workbook using the enabled cache options
            var workbookEnabled = new Workbook(filePath, loadOptionsEnabled);

            // Measure memory after loading
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterEnabled = GC.GetTotalMemory(true);
            long memoryUsedEnabled = afterEnabled - beforeEnabled;

            // -------------------------------------------------
            // Load with memory optimization (cache parsing disabled)
            // -------------------------------------------------
            var loadOptionsDisabled = new LoadOptions(LoadFormat.Xlsx)
            {
                // Reduce memory usage by disabling internal caching
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Clean up before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeDisabled = GC.GetTotalMemory(true);

            // Load the workbook using the disabled cache options
            var workbookDisabled = new Workbook(filePath, loadOptionsDisabled);

            // Measure memory after loading
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterDisabled = GC.GetTotalMemory(true);
            long memoryUsedDisabled = afterDisabled - beforeDisabled;

            // Output the comparison results
            Console.WriteLine($"Memory used with cache parsing enabled : {memoryUsedEnabled / 1024} KB");
            Console.WriteLine($"Memory used with cache parsing disabled: {memoryUsedDisabled / 1024} KB");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during loading or measurement
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
