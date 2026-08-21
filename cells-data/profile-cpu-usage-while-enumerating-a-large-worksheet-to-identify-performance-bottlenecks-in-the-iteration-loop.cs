// Title: Profile CPU and detect slow cells while enumerating a large worksheet with Aspose.Cells for .NET
// Description: Creates a 10,000 × 50 worksheet, enables CellsData access cache, iterates through every cell with an IEnumerator, measures total and per‑cell execution time using Stopwatch, logs cells exceeding a 1 ms threshold, closes the cache, and saves the file.
// Keywords: Aspose.Cells | C# | cell enumeration performance | CPU profiling | StartAccessCache | large worksheet iteration | per‑cell timing | performance bottleneck detection
// Common Searches: Aspose.Cells enumerate cells performance .NET | measure cell iteration time with Stopwatch | StartAccessCache effect on worksheet traversal speed | log slow cells during Aspose.Cells enumeration | profile CPU usage while reading large Excel file
// Developer Intent: Measure CPU consumption and isolate cells that cause delays when iterating over a massive worksheet with Aspose.Cells.
// Use Cases: Identify individual cells that exceed a processing‑time threshold for optimization. | Compare enumeration speed with and without CellsData access cache to choose the best strategy. | Gather benchmark data for total and per‑cell execution time in high‑volume spreadsheet processing.
// AI Prompts: Rewrite the enumeration loop to capture both CPU and memory metrics per cell using Aspose.Cells. | Suggest alternative iteration methods (e.g., Range, DataTable) and compare their performance with IEnumerator. | Generate a performance report template that aggregates per‑cell timings and highlights the top 10 slowest cells.

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

// Creates a 10,000 × 50 worksheet, enables CellsData access cache, iterates through every cell with an IEnumerator, measures total and per‑cell execution time using Stopwatch, logs cells exceeding a 1 ms threshold, closes the cache, and saves the file.
class CpuProfilingEnumeration
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a large worksheet (e.g., 10,000 rows × 50 columns)
        int totalRows = 10000;
        int totalCols = 50;
        for (int r = 0; r < totalRows; r++)
        {
            for (int c = 0; c < totalCols; c++)
            {
                cells[r, c].PutValue(r * totalCols + c);
            }
        }

        // Start access cache for cell data to improve read‑only enumeration performance
        workbook.StartAccessCache(AccessCacheOptions.CellsData);

        // Get the cells enumerator
        IEnumerator enumerator = cells.GetEnumerator();

        // Stopwatch to measure total enumeration time
        Stopwatch totalTimer = Stopwatch.StartNew();

        // Stopwatch to measure time spent on each individual cell (optional)
        Stopwatch cellTimer = new Stopwatch();

        long processedCells = 0;

        while (enumerator.MoveNext())
        {
            cellTimer.Restart();

            // Current cell from the enumerator
            Cell cell = (Cell)enumerator.Current;

            // Example processing: read the cell value
            var value = cell.Value;

            cellTimer.Stop();

            // Log cells that take unusually long to process (e.g., > 1 ms)
            if (cellTimer.ElapsedMilliseconds > 1)
            {
                Console.WriteLine($"Slow cell {cell.Name}: {cellTimer.ElapsedMilliseconds} ms");
            }

            processedCells++;
        }

        totalTimer.Stop();

        // Close the cache after enumeration is finished
        workbook.CloseAccessCache(AccessCacheOptions.CellsData);

        Console.WriteLine($"Enumerated {processedCells} cells in {totalTimer.Elapsed.TotalSeconds:F2} seconds.");

        // Save the workbook (optional, demonstrates lifecycle rule usage)
        workbook.Save("LargeWorksheet.xlsx");
    }
}
