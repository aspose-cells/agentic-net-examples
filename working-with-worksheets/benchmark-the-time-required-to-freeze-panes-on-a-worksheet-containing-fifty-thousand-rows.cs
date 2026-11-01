// Title: Measure the time to apply FreezePanes on a worksheet with 50,000 rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook with 50,000 rows, applies FreezePanes to the first row, and prints the elapsed milliseconds. | Generate a .NET snippet that populates a large worksheet, times the FreezePanes call, and saves the workbook for verification.
// Common Searches: asp.net benchmark freeze panes on large worksheet 50000 rows | how to measure freeze panes execution time with Aspose.Cells C# | performance of FreezePanes method on 50k rows Excel file | time taken to freeze top row in a 50,000‑row spreadsheet using Aspose.Cells
// Tags: Aspose.Cells FreezePanes performance | benchmark FreezePanes .NET | large worksheet freeze panes timing | measure Excel freeze panes latency | populate 50000 rows Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// Creates a new workbook, fills column A with 50,000 rows, uses Stopwatch to time the FreezePanes(1,0,1,0) call, prints the elapsed milliseconds, and saves the file as FreezePanesBenchmark.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with 50,000 rows of sample data
            for (int row = 0; row < 50000; row++)
            {
                // Fill column A with the row index (or any data you prefer)
                sheet.Cells[row, 0].PutValue(row + 1);
            }

            // Benchmark the time required to apply freeze panes
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Freeze the top row (row index 1 means the first row is frozen)
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            sw.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"Time to freeze panes on a sheet with 50,000 rows: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (optional, ensures the operation is fully applied)
            workbook.Save("FreezePanesBenchmark.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
