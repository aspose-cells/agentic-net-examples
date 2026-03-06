using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // -----------------------------------------------------------------
            // Caution: Use LoadOptions with MemoryPreference to reduce memory usage
            // when loading large workbooks. This setting tells Aspose.Cells to
            // keep data in a memory‑efficient mode during loading.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Open the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // -----------------------------------------------------------------
            // Additional caution: Set the workbook's default memory setting.
            // This affects worksheets created after the workbook is loaded.
            // -----------------------------------------------------------------
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Example operation: read the first worksheet name
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + firstSheet.Name);

            // -----------------------------------------------------------------
            // Optional performance boost: start an access cache.
            // This caches style and formatting information for faster reads.
            // Remember to close the cache after the operation.
            // -----------------------------------------------------------------
            workbook.StartAccessCache(AccessCacheOptions.All);

            // Perform a lightweight read operation while the cache is active
            for (int row = 0; row < Math.Min(10, firstSheet.Cells.MaxDataRow + 1); row++)
            {
                Cell cell = firstSheet.Cells[row, 0];
                Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
            }

            // Close the access cache to release resources
            workbook.CloseAccessCache(AccessCacheOptions.All);

            // Save the workbook to a new file (memory‑optimized settings are retained)
            string outputPath = "output_optimized.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}' with memory‑optimized settings.");

            // Dispose the workbook to free unmanaged resources
            workbook.Dispose();
        }
    }
}