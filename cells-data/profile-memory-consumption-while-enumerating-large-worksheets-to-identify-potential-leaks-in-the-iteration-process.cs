using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

class MemoryProfilingExample
{
    static void Main()
    {
        // Load a large workbook (replace with actual file path)
        string inputPath = "LargeWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Use FileCache mode to keep memory usage low during iteration
        cells.MemorySetting = MemorySetting.FileCache;

        // Force a full garbage collection before starting measurements
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        long initialMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Initial memory: {initialMemory / 1024 / 1024} MB");

        // Enumerate rows sequentially (recommended for FileCache mode)
        IEnumerator rowEnumerator = cells.Rows.GetEnumerator();
        int processedRows = 0;
        Stopwatch sw = Stopwatch.StartNew();

        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;

            // Enumerate cells within the current row (optional, just to access data)
            IEnumerator cellEnumerator = row.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;
                // Access the cell value to ensure the cell is actually read
                var _ = cell.Value;
            }

            processedRows++;

            // Report memory usage every 1000 rows
            if (processedRows % 1000 == 0)
            {
                long currentMemory = GC.GetTotalMemory(false);
                Console.WriteLine($"Rows processed: {processedRows}, Elapsed: {sw.Elapsed.TotalSeconds:F1}s, Memory: {currentMemory / 1024 / 1024} MB");
            }
        }

        sw.Stop();
        long finalMemory = GC.GetTotalMemory(true);
        Console.WriteLine($"Finished processing. Total rows: {processedRows}, Time: {sw.Elapsed.TotalSeconds:F1}s");
        Console.WriteLine($"Final memory: {finalMemory / 1024 / 1024} MB");
        Console.WriteLine($"Memory change: {(finalMemory - initialMemory) / 1024 / 1024} MB");

        // Save the workbook (demonstrates the required save rule)
        workbook.Save("ProcessedWorkbook.xlsx");
    }
}