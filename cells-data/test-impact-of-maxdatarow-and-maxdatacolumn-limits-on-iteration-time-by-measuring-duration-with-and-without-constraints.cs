// Title: Benchmark Cells.MaxDataRow vs Cached Values – Impact on Iteration Speed in Aspose.Cells for .NET
// Description: Creates a 5,000‑row × 100‑column worksheet, fills it with numbers, then measures two loops: one that reads Cells.MaxDataRow/MaxDataColumn on every iteration and another that caches these limits first. The program outputs the elapsed milliseconds for each approach and saves the workbook, illustrating loop‑level performance differences in Aspose.Cells.
// Keywords: Aspose.Cells performance | MaxDataRow benchmark | MaxDataColumn caching | C# worksheet iteration speed | loop optimization Aspose.Cells | cell access overhead .NET | large worksheet processing | benchmarking Aspose.Cells | iteration time measurement | performance best practice
// Common Searches: Aspose.Cells benchmark MaxDataRow vs cached | how to improve loop performance in Aspose.Cells | measure Cells.MaxDataColumn overhead | caching MaxDataRow for faster iteration | Aspose.Cells iteration speed test | performance impact of property access in Aspose.Cells loops | optimal way to traverse large worksheets in C#
// Developer Intent: Evaluate the speed benefit of caching MaxDataRow and MaxDataColumn when iterating over a worksheet.
// Use Cases: Determine whether caching worksheet limits reduces processing time in bulk data reads. | Validate performance before implementing large‑scale data transformations with Aspose.Cells. | Provide a reference example for high‑throughput cell iteration in .NET applications.
// AI Prompts: Write C# code that compares iteration time using Cells.MaxDataRow/MaxDataColumn each loop versus cached values in Aspose.Cells. | Explain why accessing MaxDataRow inside a tight loop can degrade performance and suggest caching strategies. | Create a template for reporting benchmark results of the two iteration methods, including interpretation guidelines.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMaxDataImpact
{
    // Creates a 5,000‑row × 100‑column worksheet, fills it with numbers, then measures two loops: one that reads Cells.MaxDataRow/MaxDataColumn on every iteration and another that caches these limits first. The program outputs the elapsed milliseconds for each approach and saves the workbook, illustrating loop‑level performance differences in Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define size of test data
            const int totalRows = 5000;   // number of rows to populate
            const int totalCols = 100;    // number of columns to populate

            // Fill the worksheet with sample numeric data
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue(r * totalCols + c);
                }
            }

            // -----------------------------------------------------------------
            // Scenario 1: Use MaxDataRow/MaxDataColumn property inside the loop
            // (property is evaluated on every iteration, which is costly)
            // -----------------------------------------------------------------
            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int r = 0; r <= cells.MaxDataRow; r++)          // property accessed each iteration
            {
                for (int c = 0; c <= cells.MaxDataColumn; c++)  // property accessed each iteration
                {
                    // Access the cell value (no further processing needed for timing)
                    var _ = cells[r, c].Value;
                }
            }

            sw.Stop();
            long timeWithPropertyEachLoop = sw.ElapsedMilliseconds;

            // -----------------------------------------------------------------
            // Scenario 2: Cache MaxDataRow/MaxDataColumn before the loop
            // (avoids repeated property evaluation)
            // -----------------------------------------------------------------
            int maxRow = cells.MaxDataRow;       // cached once
            int maxCol = cells.MaxDataColumn;    // cached once

            sw.Restart();

            for (int r = 0; r <= maxRow; r++)
            {
                for (int c = 0; c <= maxCol; c++)
                {
                    var _ = cells[r, c].Value;
                }
            }

            sw.Stop();
            long timeWithCachedValues = sw.ElapsedMilliseconds;

            // Output the measured durations
            Console.WriteLine($"Iteration with property call each loop: {timeWithPropertyEachLoop} ms");
            Console.WriteLine($"Iteration with cached MaxDataRow/MaxDataColumn: {timeWithCachedValues} ms");

            // Save the workbook (optional, demonstrates normal save lifecycle)
            workbook.Save("MaxDataImpactResult.xlsx");
        }
    }
}
