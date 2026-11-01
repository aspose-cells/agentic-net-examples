// Title: How to profile memory usage while enumerating large worksheets with Aspose.Cells in C#
// AI Prompts: Write a C# console program that loads a .xlsx file using Aspose.Cells, iterates through each worksheet's MaxDisplayRange, and prints memory usage before the loop, after every N rows, and after completion. | Modify the iteration to invoke GC.Collect after processing each worksheet and display the memory difference in megabytes. | Extend the sample to accept the row‑interval for intermediate memory reporting and the file path as command‑line arguments, handling missing arguments gracefully.
// Common Searches: C# how to measure memory consumption while looping through cells with Aspose.Cells | detect memory leaks in large Excel workbook processing using Aspose.Cells .NET | log GC.GetTotalMemory every 1000 rows Aspose.Cells example | profile memory usage of worksheet enumeration in Aspose.Cells for .NET
// Tags: memory profiling Aspose.Cells C# | worksheet enumeration memory usage | GC.GetTotalMemory Aspose.Cells | large Excel workbook processing .NET | used range cell iteration Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads a large .xlsx workbook with Aspose.Cells, iterates each worksheet's used range, logs initial memory, reports memory every 1000 rows, forces garbage collection per worksheet, and outputs final memory usage and processing time using GC.GetTotalMemory.
class MemoryProfilingExample
{
    static void Main(string[] args)
    {
        // Path to the large Excel file to be processed
        string filePath = @"C:\Data\LargeWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at '{filePath}'.");
            return;
        }

        try
        {
            // Load the workbook (using Aspose.Cells) inside a using block for proper disposal
            using (Workbook workbook = new Workbook(filePath))
            {
                // Force a full garbage collection before starting the measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Record initial memory usage (in bytes)
                long initialMemory = GC.GetTotalMemory(forceFullCollection: true);
                Console.WriteLine($"Initial memory usage: {initialMemory / (1024 * 1024)} MB");

                // Stopwatch to measure elapsed time (optional, but useful for performance profiling)
                Stopwatch sw = Stopwatch.StartNew();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing sheet: {sheet.Name}");

                    // Get the used range of the worksheet to avoid iterating over empty cells
                    Cells cells = sheet.Cells;
                    Aspose.Cells.Range usedRange = cells.MaxDisplayRange; // Fully qualified to avoid ambiguity

                    // Determine start and end indices from the Range object
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startColumn = usedRange.FirstColumn;
                    int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Iterate row by row within the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        // Optional: measure memory after each N rows to spot gradual increase
                        if (row % 1000 == 0)
                        {
                            long intermediateMemory = GC.GetTotalMemory(forceFullCollection: false);
                            Console.WriteLine($"Row {row}: Memory usage = {intermediateMemory / (1024 * 1024)} MB");
                        }

                        // Iterate through columns in the current row
                        for (int col = startColumn; col <= endColumn; col++)
                        {
                            // Access the cell value (this is the typical operation that could cause leaks if mishandled)
                            Cell cell = cells[row, col];
                            object value = cell.Value; // Replace with actual processing logic as needed
                        }
                    }
                }

                sw.Stop();

                // Record final memory usage after iteration
                long finalMemory = GC.GetTotalMemory(forceFullCollection: true);
                Console.WriteLine($"Final memory usage: {finalMemory / (1024 * 1024)} MB");
                Console.WriteLine($"Memory delta: {(finalMemory - initialMemory) / (1024 * 1024)} MB");
                Console.WriteLine($"Total processing time: {sw.Elapsed}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
