// Title: Profile Memory Usage While Enumerating Large Worksheets with Aspose.Cells for .NET
// Description: C# sample that creates a 200,000‑row workbook, switches Cells.MemorySetting to FileCache, iterates rows with IEnumerator, logs managed memory at intervals, measures execution time, and disposes resources to detect possible memory leaks during large‑scale enumeration.
// Keywords: Aspose.Cells memory profiling | FileCache mode enumeration | large worksheet iteration .NET | managed memory leak detection | row enumeration performance | C# Aspose.Cells memory usage | benchmark spreadsheet processing
// Common Searches: Aspose.Cells how to profile memory while reading rows | enumerate rows in FileCache mode without leaks | measure managed memory during large worksheet processing | C# Aspose.Cells performance tips for 200k rows | detect memory growth in Aspose.Cells iteration
// Developer Intent: The developer needs to monitor and verify managed memory consumption when iterating over a massive worksheet to ensure the FileCache setting prevents leaks and to benchmark performance.
// Use Cases: Identify memory spikes in server‑side spreadsheet processing pipelines | Validate that FileCache mode keeps memory stable during row‑by‑row reads | Benchmark enumeration speed and memory impact for bulk data imports
// AI Prompts: Generate C# code that records managed memory every 10,000 rows while iterating a worksheet with Aspose.Cells FileCache mode. | Suggest optimizations to minimize memory delta during large worksheet enumeration in Aspose.Cells. | Create a unit‑test method that asserts memory usage stays below a defined threshold while processing 200,000 rows.

using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    // C# sample that creates a 200,000‑row workbook, switches Cells.MemorySetting to FileCache, iterates rows with IEnumerator, logs managed memory at intervals, measures execution time, and disposes resources to detect possible memory leaks during large‑scale enumeration.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set memory usage mode to FileCache to reduce in‑memory footprint
            cells.MemorySetting = MemorySetting.FileCache;

            // Populate a large number of rows (e.g., 200,000 rows, 5 columns)
            const int totalRows = 200_000;
            const int totalCols = 5;
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Save the workbook (uses the provided save rule)
            workbook.Save("LargeData.xlsx");

            // Prepare for profiling
            Console.WriteLine("Starting enumeration and memory profiling...");
            long initialMemory = GC.GetTotalMemory(forceFullCollection: true);
            Console.WriteLine($"Initial managed memory: {initialMemory / 1024 / 1024} MB");

            Stopwatch sw = Stopwatch.StartNew();

            // Enumerate rows sequentially (recommended for FileCache mode)
            IEnumerator rowEnum = cells.Rows.GetEnumerator();
            int processedRows = 0;
            while (rowEnum.MoveNext())
            {
                Row row = (Row)rowEnum.Current;

                // Access each cell in the row to simulate work
                IEnumerator cellEnum = row.GetEnumerator();
                while (cellEnum.MoveNext())
                {
                    Cell cell = (Cell)cellEnum.Current;
                    // Simple read operation
                    string val = cell.StringValue;
                }

                processedRows++;

                // Periodically report memory usage to detect leaks
                if (processedRows % 20_000 == 0)
                {
                    long currentMemory = GC.GetTotalMemory(forceFullCollection: true);
                    Console.WriteLine($"Rows processed: {processedRows}, Managed memory: {currentMemory / 1024 / 1024} MB");
                }
            }

            sw.Stop();
            long finalMemory = GC.GetTotalMemory(forceFullCollection: true);
            Console.WriteLine($"Enumeration completed in {sw.Elapsed.TotalSeconds:F2} seconds.");
            Console.WriteLine($"Final managed memory: {finalMemory / 1024 / 1024} MB");
            Console.WriteLine($"Memory delta: {(finalMemory - initialMemory) / 1024 / 1024} MB");

            // Dispose workbook resources (important for FileCache mode)
            workbook.Dispose();
        }
    }
}
