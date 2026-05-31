using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "LargeFile.xlsx";
            const string outputPath = "ProcessedLargeFile.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook or create a new one if the file is missing
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Reduce in‑memory footprint for large data
                workbook.Worksheets[0].Cells.MemorySetting = MemorySetting.FileCache;

                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Warm up GC and capture baseline memory usage
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                long memoryBefore = GC.GetTotalMemory(true);
                Console.WriteLine($"Memory before enumeration: {memoryBefore / 1024 / 1024} MB");

                // Start timing
                Stopwatch sw = Stopwatch.StartNew();

                // Enumerate rows sequentially (recommended for FileCache mode)
                IEnumerator rowEnum = cells.Rows.GetEnumerator();
                while (rowEnum.MoveNext())
                {
                    Row row = (Row)rowEnum.Current;

                    // Enumerate cells within the current row
                    IEnumerator cellEnum = row.GetEnumerator();
                    while (cellEnum.MoveNext())
                    {
                        Cell cell = (Cell)cellEnum.Current;
                        var value = cell.Value; // Dummy read to simulate processing
                    }
                }

                sw.Stop();

                // Measure memory after enumeration
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                long memoryAfter = GC.GetTotalMemory(true);
                Console.WriteLine($"Memory after enumeration: {memoryAfter / 1024 / 1024} MB");
                Console.WriteLine($"Memory delta: {(memoryAfter - memoryBefore) / 1024 / 1024} MB");
                Console.WriteLine($"Enumeration time: {sw.ElapsedMilliseconds} ms");

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                if (workbook != null)
                {
                    workbook.Dispose();
                }

                // Final memory check after disposal
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                long memoryFinal = GC.GetTotalMemory(true);
                Console.WriteLine($"Memory after disposal: {memoryFinal / 1024 / 1024} MB");
            }
        }
    }
}