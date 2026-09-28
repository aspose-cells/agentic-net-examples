// Title: How to profile memory usage while iterating over used cells in large Excel worksheets with Aspose.Cells for .NET
// AI Prompts: Create a C# console application that loads an Excel workbook using Aspose.Cells, iterates through the used range of every worksheet, and logs memory usage to the console after a configurable number of rows. | Adjust the enumeration code to capture a memory snapshot after each worksheet finishes processing and display the total memory delta compared with the baseline. | Add a configurable memory‑threshold check that aborts the iteration and reports the worksheet name when the memory increase exceeds the specified limit.
// Common Searches: aspnet memory profiling while reading large Excel files with Aspose.Cells | how to detect memory leaks during worksheet iteration in C# using Aspose.Cells | measure heap growth when enumerating used cells of an Excel workbook in .NET | C# code to log memory usage every 1000 rows while processing large worksheets with Aspose.Cells
// Tags: memory profiling Aspose.Cells cell enumeration | heap usage monitoring worksheet iteration .NET | used range iteration performance Aspose.Cells | periodic GC snapshot large Excel processing | memory leak detection Aspose.Cells workbook traversal

using System;
using System.Diagnostics;
using Aspose.Cells;

// Demonstrates how to measure memory consumption while walking through the used cells of each worksheet in a large Excel workbook with Aspose.Cells, reporting baseline, periodic (e.g., every 1,000 rows) and final memory usage.
class Program
{
    static void Main(string[] args)
    {
        // Path to the large workbook to be analyzed.
        string inputPath = "largeWorkbook.xlsx"; // TODO: replace with actual file path.

        // Load the workbook using Aspose.Cells.
        Workbook workbook = new Workbook(inputPath);

        // Force a full garbage collection and capture the baseline memory usage.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long baselineMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Baseline memory: {baselineMemory / 1024.0 / 1024.0:F2} MB");

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Processing worksheet: {sheet.Name}");

            // Obtain the used range to limit iteration to populated cells.
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate rows.
            for (int row = 0; row <= maxRow; row++)
            {
                // Iterate columns within the current row.
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Access the cell value. This forces Aspose.Cells to materialize the cell object.
                    object value = cells[row, col].Value;
                    // (Optional) Process the value here if needed.
                }

                // Periodically report memory usage to spot leaks during long iterations.
                if (row % 1000 == 0) // Adjust the interval based on worksheet size.
                {
                    long currentMemory = GC.GetTotalMemory(true);
                    Console.WriteLine($"Row {row}/{maxRow} - Memory: {currentMemory / 1024.0 / 1024.0:F2} MB");
                }
            }
        }

        // Final memory snapshot after full enumeration.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long finalMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Final memory: {finalMemory / 1024.0 / 1024.0:F2} MB");
        Console.WriteLine($"Memory delta: {(finalMemory - baselineMemory) / 1024.0 / 1024.0:F2} MB");
    }
}
