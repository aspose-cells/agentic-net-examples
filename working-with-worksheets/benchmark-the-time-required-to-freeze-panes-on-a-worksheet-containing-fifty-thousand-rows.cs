// Title: C# Benchmark: FreezePanes Execution Time on a 50,000‑Row Worksheet with Aspose.Cells
// Description: Creates a workbook, fills column A with 50,000 rows, measures the time to apply FreezePanes at C3 (2 rows × 2 columns), prints the elapsed milliseconds, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | FreezePanes | benchmark | performance testing | .NET | C# | large worksheet | 50,000 rows | execution time | measure latency
// Common Searches: Aspose.Cells FreezePanes performance test | how long does FreezePanes take on 50k rows C# | benchmarking FreezePanes execution time .NET | measure FreezePanes latency Aspose.Cells
// Developer Intent: Find out how many milliseconds the FreezePanes method needs on a worksheet with fifty thousand rows.
// Use Cases: Assess the impact of FreezePanes on report generation speed for massive data sets. | Compare freezing performance across different worksheet sizes to set optimal thresholds. | Validate that pane freezing meets latency requirements in automated Excel creation pipelines.
// AI Prompts: Provide a C# example that benchmarks FreezePanes on a worksheet with 100,000 rows using Aspose.Cells. | Suggest ways to reduce FreezePanes overhead when working with very large Excel files in .NET. | Explain how to capture detailed timing, including GC pauses, for the FreezePanes call in a performance test.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace FreezePanesBenchmark
{
    // Creates a workbook, fills column A with 50,000 rows, measures the time to apply FreezePanes at C3 (2 rows × 2 columns), prints the elapsed milliseconds, and saves the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate 50,000 rows with sample data
            // This ensures the worksheet has the required number of rows
            for (int row = 0; row < 50000; row++)
            {
                // Fill column A with the row number (as string)
                worksheet.Cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Benchmark the FreezePanes operation
            Stopwatch sw = Stopwatch.StartNew();

            // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
            // Using the FreezePanes(int, int, int, int) overload (rule)
            worksheet.FreezePanes(2, 2, 2, 2);

            sw.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"FreezePanes execution time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (lifecycle rule)
            workbook.Save("FreezePanesBenchmark.xlsx");
        }
    }
}
