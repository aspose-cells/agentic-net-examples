// Title: Benchmark FreezePanes Calls in Aspose.Cells for .NET Using Stopwatch
// Description: Shows how to time two different FreezePanes overloads (row/column indices and cell address) on a worksheet with System.Diagnostics.Stopwatch, output the elapsed milliseconds, and save the workbook.
// Keywords: Aspose.Cells | C# | FreezePanes | performance benchmark | stopwatch timing | execution latency | worksheet freeze | code example | GitHub | dotnet
// Common Searches: measure FreezePanes execution time Aspose.Cells | C# stopwatch FreezePanes benchmark | how long does FreezePanes take in .NET | record worksheet freeze latency Aspose | Aspose.Cells performance testing FreezePanes
// Developer Intent: Determine and log the duration of each FreezePanes operation in a .NET workbook.
// Use Cases: Compare the speed of different FreezePanes overloads when processing large spreadsheets. | Monitor the impact of repeated FreezePanes calls on overall workbook generation time. | Create performance baselines before and after applying formatting or data loading. | Export timing results to logs or analytics dashboards for continuous monitoring.
// AI Prompts: Generate a reusable method that accepts a Worksheet and FreezePanes parameters, executes the freeze, and returns the elapsed time in milliseconds. | Provide code to collect multiple FreezePanes timings and write them to a CSV file for later analysis. | Suggest best practices to minimize FreezePanes execution time when working with massive Excel files in Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsFreezePanesTiming
{
    // Shows how to time two different FreezePanes overloads (row/column indices and cell address) on a worksheet with System.Diagnostics.Stopwatch, output the elapsed milliseconds, and save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // First FreezePanes call
            Stopwatch sw1 = Stopwatch.StartNew();
            sheet.FreezePanes(3, 3, 3, 3); // Freeze at row 3, column 3 with 3 rows and 3 columns frozen
            sw1.Stop();
            Console.WriteLine($"FreezePanes call 1 elapsed time: {sw1.ElapsedMilliseconds} ms");

            // Second FreezePanes call (different position)
            Stopwatch sw2 = Stopwatch.StartNew();
            sheet.FreezePanes("E5", 5, 5); // Freeze at cell E5 with 5 rows and 5 columns frozen
            sw2.Stop();
            Console.WriteLine($"FreezePanes call 2 elapsed time: {sw2.ElapsedMilliseconds} ms");

            // Save the workbook
            workbook.Save("FreezePanesTimingDemo.xlsx");
        }
    }
}
