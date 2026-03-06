using System;
using Aspose.Cells;

namespace LargeExcelMemoryOptimized
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Optimize memory usage for large data sets
            // MemoryPreference keeps data in a compact format to reduce RAM consumption
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Optional: also set the worksheet's cells memory mode (inherits from workbook settings)
            cells.MemorySetting = MemorySetting.MemoryPreference;

            // Write a large amount of data sequentially (row by row) to benefit from the memory mode
            // Example: 1,000,000 rows with 10 columns
            const int totalRows = 1_000_000;
            const int totalCols = 10;

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Put a simple value; you can replace with any data source
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }

                // Periodically flush to avoid holding too many objects in memory
                // (not strictly required, but helps when using FileCache mode)
                if (row % 100_000 == 0)
                {
                    Console.WriteLine($"Written {row + 1} rows...");
                }
            }

            // Save the workbook to disk using the standard Save method
            workbook.Save("LargeOptimized.xlsx", SaveFormat.Xlsx);

            // Dispose the workbook to release any temporary resources (especially important for FileCache mode)
            workbook.Dispose();

            Console.WriteLine("Workbook saved successfully with memory optimization.");
        }
    }
}