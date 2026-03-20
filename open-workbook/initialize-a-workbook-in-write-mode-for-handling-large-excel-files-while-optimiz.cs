using System;
using Aspose.Cells;

namespace LargeWorkbookDemo
{
    class Program
    {
        static void Main()
        {
            // Initialize a new workbook (write mode)
            Workbook workbook = new Workbook();

            // Optimize memory usage for large files
            // MemoryPreference keeps data in a compact format to reduce RAM consumption
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

            // Example: add a large amount of data (optional, demonstrates usage)
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 100000; row++)
            {
                // Populate only the first column to keep the demo simple
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Save the workbook to disk in XLSX format
            workbook.Save("LargeWorkbook.xlsx", SaveFormat.Xlsx);

            // Clean up resources
            workbook.Dispose();

            Console.WriteLine("Workbook created and saved with optimized memory settings.");
        }
    }
}