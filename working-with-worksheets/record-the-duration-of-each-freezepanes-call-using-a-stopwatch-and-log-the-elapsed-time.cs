// Title: How to time each Worksheet.FreezePanes call with Stopwatch and log milliseconds in Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses a single Stopwatch to measure the execution time of multiple Worksheet.FreezePanes calls and prints the elapsed milliseconds. | Show how to restart a Stopwatch between successive FreezePanes operations and output each timing result to the console. | Provide an example that benchmarks FreezePanes performance, logs the timings, and saves the workbook to an .xlsx file.
// Common Searches: C# Aspose.Cells how to measure FreezePanes execution time | using Stopwatch to profile worksheet FreezePanes in .NET | log elapsed time for each FreezePanes call Aspose.Cells example | benchmarking FreezePanes performance with Aspose.Cells C# | record timing of FreezePanes rows and columns Aspose.Cells
// Tags: worksheet freeze panes timing Aspose.Cells | stopwatch benchmark freeze panes .NET | measure freeze panes performance C# | log freeze panes elapsed time Aspose.Cells | save workbook after freeze panes timing

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace FreezePanesTimingExample
{
    // The example creates a workbook, accesses the first worksheet, and uses a single Stopwatch instance to measure and log the elapsed milliseconds for three separate Worksheet.FreezePanes calls (first row, first column, and both). After timing, the workbook is saved as FreezePanesTimingResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare a stopwatch to measure each FreezePanes call
                Stopwatch stopwatch = new Stopwatch();

                // -------------------------------------------------
                // First FreezePanes call (freeze the first row)
                stopwatch.Start(); // Start timing
                // Freeze the first row (row index 1, column index 0, freeze 1 row, 0 columns)
                worksheet.FreezePanes(1, 0, 1, 0);
                stopwatch.Stop(); // Stop timing
                Console.WriteLine($"FreezePanes(1,0) elapsed: {stopwatch.ElapsedMilliseconds} ms");

                // -------------------------------------------------
                // Second FreezePanes call (freeze the first column)
                stopwatch.Restart(); // Restart timing without creating a new instance
                // Freeze the first column (row index 0, column index 1, freeze 0 rows, 1 column)
                worksheet.FreezePanes(0, 1, 0, 1);
                stopwatch.Stop();
                Console.WriteLine($"FreezePanes(0,1) elapsed: {stopwatch.ElapsedMilliseconds} ms");

                // -------------------------------------------------
                // Third FreezePanes call (freeze first row and first column)
                stopwatch.Restart();
                // Freeze both first row and first column (row index 1, column index 1, freeze 1 row, 1 column)
                worksheet.FreezePanes(1, 1, 1, 1);
                stopwatch.Stop();
                Console.WriteLine($"FreezePanes(1,1) elapsed: {stopwatch.ElapsedMilliseconds} ms");

                // Save the workbook
                string outputPath = "FreezePanesTimingResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
