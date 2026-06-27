using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Initialize a stopwatch to measure the conversion duration
            Stopwatch stopwatch = new Stopwatch();

            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Populate the workbook with sample data
            Worksheet sheet = workbook.Worksheets[0];
            for (int i = 0; i < 1000; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i}");
                sheet.Cells[i, 1].PutValue(i);
            }

            // Start measuring time before the conversion (save rule)
            stopwatch.Start();

            // Convert the workbook to PDF format
            workbook.Save("ConversionResult.pdf", SaveFormat.Pdf);

            // Stop measuring time after the conversion
            stopwatch.Stop();

            // Log the elapsed time for performance analysis
            Console.WriteLine($"Conversion completed in {stopwatch.Elapsed.TotalMilliseconds} ms.");
        }
    }
}