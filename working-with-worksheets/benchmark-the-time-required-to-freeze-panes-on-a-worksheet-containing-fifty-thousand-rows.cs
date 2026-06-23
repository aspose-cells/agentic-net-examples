using System;
using System.Diagnostics;
using Aspose.Cells;

namespace FreezePanesBenchmark
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate 50,000 rows with sample data (simple incremental numbers)
            // This ensures the worksheet has the required number of rows.
            for (int row = 0; row < 50000; row++)
            {
                // Fill column A with the row index
                sheet.Cells[row, 0].PutValue(row + 1);
            }

            // Prepare a stopwatch to measure the FreezePanes operation
            Stopwatch sw = new Stopwatch();

            // Start timing
            sw.Start();

            // Freeze panes at row 2, column 2 with 2 frozen rows and 2 frozen columns
            // (FreezePanes method rule)
            sheet.FreezePanes(2, 2, 2, 2);

            // Stop timing
            sw.Stop();

            // Output the elapsed time in milliseconds
            Console.WriteLine($"Time taken to freeze panes: {sw.ElapsedMilliseconds} ms");

            // Save the workbook (save rule)
            workbook.Save("FreezePanesBenchmark.xlsx");
        }
    }
}