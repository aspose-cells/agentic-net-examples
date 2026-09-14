// Title: How to profile CPU time while enumerating a 100,000‑row worksheet with Aspose.Cells GetEnumerator in C#
// AI Prompts: Create a workbook with 100,000 rows, enable AccessCacheOptions.CellsData, iterate through every cell using the Cells enumerator, and return the elapsed CPU time measured by Stopwatch. | Add logging that records both the total processed cell count and the elapsed seconds during large‑worksheet iteration with Aspose.Cells, then suggest where to insert additional performance counters. | Extend the example to capture peak memory usage together with CPU timing while enumerating cells, and output a concise performance summary.
// Common Searches: c# how to measure CPU time for Aspose.Cells cell enumeration on large worksheets | profiling performance of Cells.GetEnumerator in Aspose.Cells with large Excel files | using AccessCacheOptions.CellsData to speed up iteration over 100k rows in Aspose.Cells | detecting bottlenecks when iterating over millions of cells with Aspose.Cells C#
// Tags: Aspose.Cells iteration speed | C# Stopwatch profiling Aspose.Cells | AccessCacheOptions.CellsData caching | large worksheet performance optimization | CPU usage measurement Aspose.Cells

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceProfiling
{
    // The sample builds a workbook containing 100,000 rows and 10 columns, starts a cell‑data cache, enumerates all instantiated cells with Cells.GetEnumerator while timing the loop using Stopwatch, closes the cache, and prints the number of processed cells along with the elapsed seconds.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large worksheet (e.g., 100,000 rows × 10 columns)
            const int totalRows = 100000;
            const int totalCols = 10;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    cells[row, col].PutValue(row * totalCols + col);
                }
            }

            // Enable cache for cell data access to reduce overhead during enumeration
            workbook.StartAccessCache(AccessCacheOptions.CellsData);

            // Measure CPU time spent iterating through all instantiated cells
            Stopwatch sw = Stopwatch.StartNew();

            // Use the Cells.GetEnumerator method as documented
            IEnumerator enumerator = cells.GetEnumerator();
            long processedCount = 0;
            while (enumerator.MoveNext())
            {
                // Cast to Cell; accessing Value forces a read which is part of the workload
                Cell cell = (Cell)enumerator.Current;
                var value = cell.Value; // Simulate processing
                processedCount++;
            }

            sw.Stop();

            // Close the cache after the operation
            workbook.CloseAccessCache(AccessCacheOptions.CellsData);

            // Output profiling results
            Console.WriteLine($"Processed {processedCount:N0} cells.");
            Console.WriteLine($"Elapsed time: {sw.Elapsed.TotalSeconds:F2} seconds.");

            // Save the workbook (optional, demonstrates create/save rule usage)
            workbook.Save("LargeWorksheetProfiled.xlsx");
        }
    }
}
