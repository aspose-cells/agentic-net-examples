// Title: C# Memory Profiling of FreezePanes on 100 Worksheets with Aspose.Cells
// Description: A console app that creates a 100‑sheet workbook, forces garbage collection, records baseline private memory, applies FreezePanes at cell C3 on each sheet, logs memory after each call, calculates per‑sheet deltas, and saves the file. Use it to detect memory leaks or performance regressions in Aspose.Cells for .NET.
// Keywords: Aspose.Cells memory profiling | FreezePanes performance .NET | detect memory leak Aspose.Cells | private memory usage C# | benchmark FreezePanes | large workbook memory consumption | GC profiling Aspose.Cells
// Common Searches: profile memory while freezing panes Aspose.Cells | detect memory leak after FreezePanes in .NET | measure private memory growth with Aspose.Cells | how to benchmark FreezePanes performance | C# memory usage test for large workbooks
// Developer Intent: The developer wants to monitor private memory before and after repeatedly calling FreezePanes on many worksheets to identify potential memory leaks or performance issues.
// Use Cases: Validate that FreezePanes does not cause a memory leak in large workbooks. | Compare memory footprints of different FreezePanes configurations. | Create reproducible memory‑usage reports for performance testing. | Integrate memory profiling into automated CI pipelines for Aspose.Cells projects.
// AI Prompts: Write a C# method that profiles memory before and after applying FreezePanes to a configurable number of worksheets using Aspose.Cells. | Suggest techniques to reduce noise in memory measurements when profiling FreezePanes calls. | Generate code to export per‑worksheet memory deltas to a CSV file for further analysis.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace FreezePanesMemoryProfiling
{
    // A console app that creates a 100‑sheet workbook, forces garbage collection, records baseline private memory, applies FreezePanes at cell C3 on each sheet, logs memory after each call, calculates per‑sheet deltas, and saves the file. Use it to detect memory leaks or performance regressions in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure we have at least one worksheet (the default one)
            // Add additional worksheets to reach a total of 100
            for (int i = workbook.Worksheets.Count; i < 100; i++)
            {
                workbook.Worksheets.Add();
            }

            // List to hold memory usage after each FreezePanes call
            List<long> memoryUsages = new List<long>();

            // Force garbage collection before starting the test
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            // Record baseline memory
            long baseline = Process.GetCurrentProcess().PrivateMemorySize64;
            Console.WriteLine($"Baseline memory: {baseline / 1024 / 1024} MB");

            // Freeze panes on each worksheet and record memory usage
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Example: freeze at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
                sheet.FreezePanes(2, 2, 2, 2);

                // Optional: force a short GC to get more consistent measurements
                GC.Collect();
                GC.WaitForPendingFinalizers();

                long currentMemory = Process.GetCurrentProcess().PrivateMemorySize64;
                memoryUsages.Add(currentMemory);
                Console.WriteLine($"Worksheet {i + 1}: Memory = {currentMemory / 1024 / 1024} MB");
            }

            // Analyze memory growth
            Console.WriteLine("\nMemory usage delta per worksheet:");
            for (int i = 0; i < memoryUsages.Count; i++)
            {
                long delta = memoryUsages[i] - baseline;
                Console.WriteLine($"Worksheet {i + 1}: +{delta / 1024 / 1024} MB");
            }

            // Save the workbook (using the standard Save method)
            workbook.Save("FreezePanesMemoryProfile.xlsx");
            Console.WriteLine("\nWorkbook saved as FreezePanesMemoryProfile.xlsx");
        }
    }
}
