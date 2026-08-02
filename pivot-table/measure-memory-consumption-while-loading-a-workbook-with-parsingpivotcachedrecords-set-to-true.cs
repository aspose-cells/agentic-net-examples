using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be loaded
            string workbookPath = "example.xlsx";

            // Measure memory before loading
            long memoryBefore = Process.GetCurrentProcess().PrivateMemorySize64;

            // Create load options and enable parsing of pivot cached records
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingPivotCachedRecords = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Measure memory after loading
            long memoryAfter = Process.GetCurrentProcess().PrivateMemorySize64;

            // Calculate and display the memory consumption
            long memoryConsumed = memoryAfter - memoryBefore;
            Console.WriteLine($"Memory before loading: {memoryBefore / 1024 / 1024} MB");
            Console.WriteLine($"Memory after loading : {memoryAfter / 1024 / 1024} MB");
            Console.WriteLine($"Memory consumed by loading workbook with ParsingPivotCachedRecords=true: {memoryConsumed / 1024 / 1024} MB");
        }
    }
}