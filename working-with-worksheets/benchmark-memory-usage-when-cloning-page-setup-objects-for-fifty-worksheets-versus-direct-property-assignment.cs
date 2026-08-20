// Title: Aspose.Cells .NET: Memory Benchmark – PageSetup.Clone vs Direct Property Assignment for 50 Worksheets
// Description: A C# console app that creates a workbook, configures a sample PageSetup, and measures managed heap memory when the same settings are applied to 49 additional worksheets using either PageSetup.Copy (clone) or manual property assignment. The program reports the memory delta for each approach, helping developers choose the most efficient method.
// Keywords: Aspose.Cells | C# | .NET | PageSetup | memory benchmark | clone vs direct assignment | Copy method performance | worksheet page settings | memory usage measurement | Excel report optimization
// Common Searches: Aspose.Cells PageSetup memory usage benchmark | Clone PageSetup vs manual property setting performance | How to measure memory impact of PageSetup.Copy in .NET | Best way to apply identical page settings to many worksheets | Memory efficient worksheet page setup Aspose.Cells
// Developer Intent: Identify which technique—PageSetup.Copy (clone) or explicit property assignment—consumes less memory when applied to 50 worksheets in Aspose.Cells for .NET.
// Use Cases: Select the most memory‑efficient strategy for applying identical page layouts in large, server‑side workbook generation. | Optimize Excel reporting pipelines by basing implementation on concrete memory‑consumption data. | Validate that cloning PageSetup does not cause unexpected memory growth in high‑volume spreadsheet processing.
// AI Prompts: Generate a C# routine that runs the benchmark repeatedly and returns average memory usage for both cloning and direct assignment methods. | Suggest refactorings to lower memory consumption when copying PageSetup across many worksheets in Aspose.Cells. | Create a unit test that asserts the clone approach uses no more than a specified percentage of additional memory compared to manual property assignment for a given worksheet count.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryBenchmark
{
    // A C# console app that creates a workbook, configures a sample PageSetup, and measures managed heap memory when the same settings are applied to 49 additional worksheets using either PageSetup.Copy (clone) or manual property assignment. The program reports the memory delta for each approach, helping developers choose the most efficient method.
    class Program
    {
        // Configure a sample PageSetup with several properties
        static void ConfigurePageSetup(PageSetup ps)
        {
            ps.PaperSize = PaperSizeType.PaperA4;
            ps.Orientation = PageOrientationType.Landscape;
            ps.FitToPagesWide = 1;
            ps.FitToPagesTall = 0; // let height adjust automatically
            ps.PrintArea = "A1:D50";
            ps.CenterHorizontally = true;
            ps.CenterVertically = true;
        }

        // Clone PageSetup using the Copy method
        static long BenchmarkClone()
        {
            try
            {
                // Force garbage collection before measurement
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long beforeMemory = GC.GetTotalMemory(true);

                // Create workbook and source worksheet
                Workbook wb = new Workbook();
                Worksheet sourceSheet = wb.Worksheets[0];
                ConfigurePageSetup(sourceSheet.PageSetup);
                PageSetup sourceSetup = sourceSheet.PageSetup;

                // Add 49 more worksheets and clone the PageSetup
                for (int i = 1; i < 50; i++)
                {
                    int newIndex = wb.Worksheets.Add();
                    Worksheet ws = wb.Worksheets[newIndex];
                    ws.PageSetup.Copy(sourceSetup, new CopyOptions());
                }

                // Optional: save to ensure workbook is fully built
                wb.Save("CloneBenchmark.xlsx");

                // Measure memory after operation
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long afterMemory = GC.GetTotalMemory(true);
                return afterMemory - beforeMemory;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in BenchmarkClone: {ex.Message}");
                return -1;
            }
        }

        // Directly assign each property without using Copy
        static long BenchmarkDirect()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long beforeMemory = GC.GetTotalMemory(true);

                Workbook wb = new Workbook();

                // First worksheet: set properties directly
                Worksheet first = wb.Worksheets[0];
                ConfigurePageSetup(first.PageSetup);

                // Remaining worksheets: assign properties one by one
                for (int i = 1; i < 50; i++)
                {
                    int newIndex = wb.Worksheets.Add();
                    Worksheet ws = wb.Worksheets[newIndex];
                    PageSetup ps = ws.PageSetup;
                    ps.PaperSize = PaperSizeType.PaperA4;
                    ps.Orientation = PageOrientationType.Landscape;
                    ps.FitToPagesWide = 1;
                    ps.FitToPagesTall = 0;
                    ps.PrintArea = "A1:D50";
                    ps.CenterHorizontally = true;
                    ps.CenterVertically = true;
                }

                wb.Save("DirectBenchmark.xlsx");

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                long afterMemory = GC.GetTotalMemory(true);
                return afterMemory - beforeMemory;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in BenchmarkDirect: {ex.Message}");
                return -1;
            }
        }

        static void Main()
        {
            // Benchmark cloning approach
            long cloneMemory = BenchmarkClone();
            if (cloneMemory >= 0)
                Console.WriteLine($"Memory used when cloning PageSetup for 50 worksheets: {cloneMemory / 1024} KB");

            // Benchmark direct assignment approach
            long directMemory = BenchmarkDirect();
            if (directMemory >= 0)
                Console.WriteLine($"Memory used when assigning PageSetup properties directly for 50 worksheets: {directMemory / 1024} KB");
        }
    }
}
