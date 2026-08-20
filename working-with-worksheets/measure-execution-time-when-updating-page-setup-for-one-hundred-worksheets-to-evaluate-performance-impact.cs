// Title: C# – Measure PageSetup Update Time for 100 Worksheets with Aspose.Cells
// Description: Creates a workbook, adds 100 worksheets, sets FitToPagesWide = 1 and FitToPagesTall = 0 for each sheet while timing the loop with Stopwatch, prints the elapsed milliseconds, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | PageSetup performance | Stopwatch benchmark | worksheet loop timing | FitToPagesWide | FitToPagesTall | large workbook speed | API execution time
// Common Searches: Aspose.Cells how long to set PageSetup for many sheets | benchmark worksheet PageSetup changes .NET | measure performance of FitToPagesWide update | time taken to modify PageSetup across 100 worksheets | Aspose.Cells page setup speed test
// Developer Intent: Find out how many milliseconds it takes to apply PageSetup settings to a hundred worksheets.
// Use Cases: Validate that page‑setup modifications won’t bottleneck batch report generation. | Compare sequential versus optimized (e.g., parallel) updates for large workbooks. | Establish a baseline before implementing custom page‑layout logic in enterprise solutions.
// AI Prompts: Generate a C# example that measures PageSetup update time for 200 worksheets using Aspose.Cells. | Suggest ways to reduce the elapsed time when changing PageSetup properties on many sheets. | Show how to apply Parallel.ForEach to update PageSetup for a large worksheet collection and capture the duration. | Create a performance test that logs both CPU and memory usage while modifying PageSetup across 500 worksheets.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // Creates a workbook, adds 100 worksheets, sets FitToPagesWide = 1 and FitToPagesTall = 0 for each sheet while timing the loop with Stopwatch, prints the elapsed milliseconds, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure the workbook has 100 worksheets
            // The first worksheet already exists, so add 99 more
            for (int i = 1; i < 100; i++)
            {
                workbook.Worksheets.Add();
            }

            // Prepare a stopwatch to measure the time taken to update page setup
            Stopwatch sw = new Stopwatch();

            // Start timing
            sw.Start();

            // Update a page setup property for each worksheet
            // Here we set FitToPagesWide = 1 for demonstration
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.FitToPagesWide = 1;
                // Optionally, also set FitToPagesTall = 0 to let height adjust automatically
                sheet.PageSetup.FitToPagesTall = 0;
            }

            // Stop timing
            sw.Stop();

            // Output the elapsed time
            Console.WriteLine($"Time taken to update page setup for 100 worksheets: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, demonstrates usage of the save API)
            workbook.Save("PageSetupPerformance.xlsx");
        }
    }
}
