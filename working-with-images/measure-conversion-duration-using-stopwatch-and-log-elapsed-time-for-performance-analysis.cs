using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Populate some data to simulate a realistic conversion workload
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 5000; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Initialize Stopwatch to measure conversion duration
            Stopwatch sw = new Stopwatch();

            // Start timing just before the conversion (save) operation
            sw.Start();

            // Perform the conversion: save the workbook to a different format (e.g., PDF)
            // This uses the standard Save method (save rule)
            workbook.Save("ConversionResult.pdf", SaveFormat.Pdf);

            // Stop timing after the operation completes
            sw.Stop();

            // Log the elapsed time for performance analysis
            Console.WriteLine($"Conversion completed in {sw.Elapsed.TotalMilliseconds} ms.");
        }
    }
}