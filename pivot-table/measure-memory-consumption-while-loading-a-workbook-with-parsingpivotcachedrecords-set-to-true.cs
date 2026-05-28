using System;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be loaded
            string workbookPath = "example.xlsx";

            // Create load options and enable parsing of pivot cached records
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingPivotCachedRecords = true;

            // Force a garbage collection and get the memory usage before loading
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryBefore = GC.GetTotalMemory(true);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Force a garbage collection and get the memory usage after loading
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryAfter = GC.GetTotalMemory(true);

            // Calculate and display the memory consumed by the load operation
            long memoryConsumed = memoryAfter - memoryBefore;
            Console.WriteLine($"Memory before loading: {memoryBefore:N0} bytes");
            Console.WriteLine($"Memory after loading : {memoryAfter:N0} bytes");
            Console.WriteLine($"Memory consumed while loading (ParsingPivotCachedRecords = true): {memoryConsumed:N0} bytes");

            // Optional: Save the workbook to verify it was loaded correctly
            workbook.Save("LoadedWithPivotCache.xlsx", SaveFormat.Xlsx);
        }
    }
}