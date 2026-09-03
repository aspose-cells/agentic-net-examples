// Title: Measure the time taken by Worksheet.FreezePanes on a 50,000‑row worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that fills a worksheet with 50,000 rows of data and uses Stopwatch to record the milliseconds required for Worksheet.FreezePanes to lock the first row. | Create a console program that initializes an Aspose.Cells workbook, populates column A with 50,000 numeric values, applies FreezePanes(1,0,1,0), and prints the elapsed time. | Write a C# example that measures both the FreezePanes operation and the total execution time for a large worksheet using Aspose.Cells and System.Diagnostics.Stopwatch.
// Common Searches: how long does Worksheet.FreezePanes take on a sheet with 50000 rows in Aspose.Cells | Aspose.Cells performance test for freezing top row in large Excel file | C# benchmark FreezePanes method with 50k rows using Stopwatch | measure freeze panes execution time Aspose.Cells .NET large worksheet | profiling Worksheet.FreezePanes performance in Aspose.Cells
// Tags: Aspose.Cells FreezePanes timing | C# large worksheet freeze operation | Excel row freeze performance .NET | Worksheet.FreezePanes latency measurement | Aspose.Cells workbook speed test

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a new workbook, fills column A with 50,000 numbers, applies Worksheet.FreezePanes to lock the first row, and uses Stopwatch to output the milliseconds required for the freeze operation and the total execution time.
class FreezePanesBenchmark
{
    static void Main()
    {
        try
        {
            // Measure total execution time (optional)
            Stopwatch totalStopwatch = Stopwatch.StartNew();

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate 50,000 rows with sample data (column A)
            for (int row = 0; row < 50000; row++)
            {
                sheet.Cells[row, 0].PutValue(row + 1);
            }

            // Benchmark the FreezePanes operation
            Stopwatch freezeStopwatch = Stopwatch.StartNew();

            // Freeze the top row (row index 1) and no columns.
            // Parameters: firstRow (first unfrozen row), firstColumn (first unfrozen column),
            // totalRows (rows to freeze), totalColumns (columns to freeze)
            sheet.FreezePanes(1, 0, 1, 0);

            freezeStopwatch.Stop();

            Console.WriteLine($"Time to freeze panes on a sheet with 50,000 rows: {freezeStopwatch.Elapsed.TotalMilliseconds} ms");

            // Optional: Save the workbook to verify the freeze (not required for benchmarking)
            // workbook.Save("FreezePanesResult.xlsx");

            totalStopwatch.Stop();
            Console.WriteLine($"Total execution time: {totalStopwatch.Elapsed.TotalMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
