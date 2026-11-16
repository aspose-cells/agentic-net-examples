// Title: Measure execution time of Worksheet.FreezePanes on a 50,000‑row Excel workbook using Aspose.Cells for .NET
// Common Searches: C# how to time Worksheet.FreezePanes on a large Excel file with Aspose.Cells | benchmark freeze panes performance for 50000 rows using Aspose.Cells .NET | profile execution speed of FreezePanes method in Aspose.Cells workbook | measure latency of freezing top row and first column in a 50k row sheet | Aspose.Cells performance test for FreezePanes on big worksheets
// Tags: Worksheet.FreezePanes execution timing | Aspose.Cells large worksheet performance | C# Excel freeze panes latency measurement | benchmarking Aspose.Cells worksheet operations | performance profiling of Excel freeze panes in .NET

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace FreezePanesBenchmarkApp
{
    // The sample creates a new workbook, fills column A with 50,000 rows of text, starts a Stopwatch, calls Worksheet.FreezePanes to lock the first row and column, stops the timer, prints the elapsed milliseconds, and saves the file as FreezePanesBenchmark.xlsx.
    class FreezePanesBenchmark
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with 50,000 rows of sample data
                for (int row = 0; row < 50000; row++)
                {
                    // Fill column A with a simple text value
                    sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                }

                // Start timing the freeze panes operation
                Stopwatch sw = Stopwatch.StartNew();

                // Freeze the top row (row 1) and first column (column A)
                // Parameters: first scrollable row (0‑based), first scrollable column,
                // number of rows to freeze, number of columns to freeze
                sheet.FreezePanes(1, 1, 1, 1);

                // Stop timing
                sw.Stop();

                // Output the elapsed time in milliseconds
                Console.WriteLine($"Time to freeze panes: {sw.ElapsedMilliseconds} ms");

                // Define output file path
                string outputPath = "FreezePanesBenchmark.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
