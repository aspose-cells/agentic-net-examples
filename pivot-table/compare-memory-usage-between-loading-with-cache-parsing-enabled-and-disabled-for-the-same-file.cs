// Title: Measure and compare memory usage of loading an XLSX workbook with Aspose.Cells access cache enabled vs disabled in C#
// AI Prompts: Create a C# console program that loads an Excel file with Aspose.Cells, captures memory before and after the load, then repeats the process with StartAccessCache enabled and prints both memory results. | Write C# code that disposes the first Workbook, forces garbage collection, reloads the same file, starts AccessCacheOptions.All, reads a cell, closes the cache, and outputs the memory delta. | Generate a C# snippet that benchmarks memory consumption of Aspose.Cells Workbook loading with and without the access cache, using GC.GetTotalMemory and proper cleanup for accurate profiling.
// Common Searches: how to benchmark Aspose.Cells memory usage when enabling access cache in C# | C# compare memory footprint of Workbook loading with StartAccessCache versus normal load | measure memory impact of Aspose.Cells access cache for large XLSX files | Aspose.Cells memory profiling example using GC.GetTotalMemory and access cache
// Tags: Aspose.Cells memory profiling workbook load | StartAccessCache memory impact | Workbook loading memory usage C# | access cache performance Aspose.Cells | GC.GetTotalMemory Aspose.Cells benchmark

using System;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    // The example loads the same XLSX file twice—first without the access cache and then with StartAccessCache/CloseAccessCache—while measuring memory before and after each load using GC.GetTotalMemory, allowing a direct comparison of memory consumption.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that will be used for both tests
            string filePath = "sample.xlsx";

            // -------------------------------------------------
            // Test 1: Load workbook without using access cache
            // -------------------------------------------------
            // Force garbage collection to get a clean baseline
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Memory before loading
            long memBeforeNoCache = GC.GetTotalMemory(true);

            // Load the workbook normally
            Workbook wbNoCache = new Workbook(filePath);

            // Memory after loading (no cache started)
            long memAfterNoCache = GC.GetTotalMemory(true);
            long usedNoCache = memAfterNoCache - memBeforeNoCache;

            Console.WriteLine($"Memory used without access cache: {usedNoCache} bytes");

            // -------------------------------------------------
            // Test 2: Load workbook and enable access cache
            // -------------------------------------------------
            // Clean up previous workbook and force GC again
            wbNoCache.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Memory before loading with cache
            long memBeforeCache = GC.GetTotalMemory(true);

            // Load the workbook
            Workbook wbCache = new Workbook(filePath);

            // Start access cache for all possible optimizations
            wbCache.StartAccessCache(AccessCacheOptions.All);

            // Perform a simple read operation to ensure the cache is exercised
            var firstCellValue = wbCache.Worksheets[0].Cells[0, 0].Value;

            // Close the cache after the operation
            wbCache.CloseAccessCache(AccessCacheOptions.All);

            // Memory after loading and using cache
            long memAfterCache = GC.GetTotalMemory(true);
            long usedCache = memAfterCache - memBeforeCache;

            Console.WriteLine($"Memory used with access cache: {usedCache} bytes");

            // Clean up
            wbCache.Dispose();
        }
    }
}
