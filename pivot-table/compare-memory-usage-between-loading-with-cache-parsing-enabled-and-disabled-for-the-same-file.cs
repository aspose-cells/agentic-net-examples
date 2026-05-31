using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded for comparison
            string filePath = "sample.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // ------------------------------------------------------------
                // 1. Load workbook without using access cache (cache parsing disabled)
                // ------------------------------------------------------------
                long memoryBeforeNoCache = Process.GetCurrentProcess().PrivateMemorySize64;

                Workbook wbNoCache = null;
                try
                {
                    wbNoCache = new Workbook(filePath);
                    long memoryAfterNoCache = Process.GetCurrentProcess().PrivateMemorySize64;
                    long memoryUsedNoCache = memoryAfterNoCache - memoryBeforeNoCache;

                    // Store result for later output
                    Console.WriteLine("Memory usage without Access Cache: {0:N0} bytes", memoryUsedNoCache);
                }
                finally
                {
                    wbNoCache?.Dispose();
                    wbNoCache = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                // ------------------------------------------------------------
                // 2. Load workbook and enable access cache (cache parsing enabled)
                // ------------------------------------------------------------
                long memoryBeforeCache = Process.GetCurrentProcess().PrivateMemorySize64;

                Workbook wbCache = null;
                long memoryUsedCache = 0;
                try
                {
                    wbCache = new Workbook(filePath);
                    // Start access cache for all possible optimizations
                    wbCache.StartAccessCache(AccessCacheOptions.All);

                    long memoryAfterCache = Process.GetCurrentProcess().PrivateMemorySize64;
                    memoryUsedCache = memoryAfterCache - memoryBeforeCache;
                }
                finally
                {
                    // Close the cache and clean up
                    wbCache?.CloseAccessCache(AccessCacheOptions.All);
                    wbCache?.Dispose();
                    wbCache = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                // ------------------------------------------------------------
                // 3. Output the comparison results
                // ------------------------------------------------------------
                Console.WriteLine("Memory usage comparison for loading \"{0}\":", filePath);
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("Without Access Cache (cache parsing disabled): {0:N0} bytes", 
                    Process.GetCurrentProcess().PrivateMemorySize64 - memoryBeforeNoCache);
                Console.WriteLine("With Access Cache (cache parsing enabled)   : {0:N0} bytes", memoryUsedCache);
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("Difference (Cache - NoCache)                : {0:N0} bytes", 
                    memoryUsedCache - (Process.GetCurrentProcess().PrivateMemorySize64 - memoryBeforeNoCache));
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}