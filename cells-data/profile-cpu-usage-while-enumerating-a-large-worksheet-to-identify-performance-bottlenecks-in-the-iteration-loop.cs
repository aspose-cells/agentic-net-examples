using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

class ProfileWorksheetEnumeration
{
    static void Main()
    {
        // Load a large workbook (replace with the actual file path)
        string inputPath = "LargeWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath); // load

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Start cache for cells data to improve read‑only access performance
        workbook.StartAccessCache(AccessCacheOptions.CellsData);

        // Stopwatch to measure total enumeration time
        Stopwatch totalTimer = Stopwatch.StartNew();

        // Get the cells enumerator
        IEnumerator enumerator = cells.GetEnumerator();

        long processedCells = 0;
        // Stopwatch to measure time for each batch (e.g., every 100,000 cells)
        Stopwatch batchTimer = Stopwatch.StartNew();

        while (enumerator.MoveNext())
        {
            // Cast the current object to Cell
            Cell cell = (Cell)enumerator.Current;

            // Example processing: read the cell value (no modification)
            var value = cell.Value;

            processedCells++;

            // Report progress every 100,000 cells
            if (processedCells % 100_000 == 0)
            {
                batchTimer.Stop();
                Console.WriteLine($"{processedCells} cells processed in {batchTimer.Elapsed.TotalSeconds:F2} seconds");
                batchTimer.Restart();
            }
        }

        totalTimer.Stop();

        Console.WriteLine($"Total cells processed: {processedCells}");
        Console.WriteLine($"Total enumeration time: {totalTimer.Elapsed.TotalSeconds:F2} seconds");

        // Close the cache after the read‑only operation is finished
        workbook.CloseAccessCache(AccessCacheOptions.CellsData);

        // Save the workbook (no changes made, just demonstrating the save lifecycle)
        workbook.Save("ProcessedWorkbook.xlsx");
    }
}