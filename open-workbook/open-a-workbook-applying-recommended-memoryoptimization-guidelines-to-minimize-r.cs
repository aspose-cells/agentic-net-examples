using System;
using Aspose.Cells;

class MemoryOptimizedWorkbookDemo
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";
        // Path where the optimized workbook will be saved
        string outputPath = "output_optimized.xlsx";

        // Create LoadOptions and set memory usage to MemoryPreference
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the workbook with the memory‑optimized load options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Apply the same memory setting to the workbook settings for any new worksheets
        workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

        // Start access cache to improve performance for bulk operations
        workbook.StartAccessCache(AccessCacheOptions.All);

        // Example operation: add a new worksheet and write some data
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
        newSheet.Name = "OptimizedSheet";
        newSheet.Cells["A1"].PutValue("Memory Optimized");

        // Close the access cache after operations are complete
        workbook.CloseAccessCache(AccessCacheOptions.All);

        // Save the workbook using the standard Save method
        workbook.Save(outputPath, SaveFormat.Xlsx);

        // Dispose the workbook to release any temporary resources (important for FileCache mode)
        workbook.Dispose();

        Console.WriteLine("Workbook processed with memory‑optimization settings.");
    }
}