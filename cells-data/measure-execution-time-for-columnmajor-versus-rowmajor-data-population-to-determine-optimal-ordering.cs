// Title: Benchmark row‑major and column‑major cell population speed with Aspose.Cells in C#
// AI Prompts: Write C# code that fills a 5000 × 200 double array into an Aspose.Cells worksheet using a row‑first loop and records the elapsed time with Stopwatch. | Add a second loop that populates the same worksheet in column‑first order, clearing the sheet before the second pass, and print both timings. | Extend the program to repeat each population method several times and output the average write time for the row‑first and column‑first approaches.
// Common Searches: Aspose.Cells C# benchmark row major vs column major write speed | how to measure cell write performance with Aspose.Cells .NET | optimal data population order for large Excel worksheets using Aspose.Cells | compare rowwise and columnwise filling performance in Aspose.Cells | stopwatch timing of Excel cell insertion Aspose.Cells C# example
// Tags: row-major cell population Aspose.Cells | column-major worksheet fill Aspose.Cells | Aspose.Cells performance benchmarking | stopwatch timing Aspose.Cells writes | large dataset Excel export .NET

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    // The sample creates a 5000 × 200 double array, writes the values to an Aspose.Cells worksheet first by iterating rows then columns, measures the duration with Stopwatch, clears the sheet, repeats the write by iterating columns then rows, measures again, prints both elapsed times, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Define dimensions of the test data
            const int totalRows = 5000;   // number of rows
            const int totalCols = 200;    // number of columns

            // Prepare a 2‑dimensional array with sample data (double values)
            double[,] data = new double[totalRows, totalCols];
            for (int r = 0; r < totalRows; r++)
                for (int c = 0; c < totalCols; c++)
                    data[r, c] = r * 0.1 + c * 0.01;   // simple deterministic value

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------- Row‑major population --------------------
            Stopwatch swRow = Stopwatch.StartNew();

            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < totalCols; c++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }

            swRow.Stop();
            Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

            // Clear the worksheet before the next test
            cells.Clear();

            // -------------------- Column‑major population --------------------
            Stopwatch swCol = Stopwatch.StartNew();

            for (int c = 0; c < totalCols; c++)
            {
                for (int r = 0; r < totalRows; r++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }

            swCol.Stop();
            Console.WriteLine($"Column‑major population time: {swCol.ElapsedMilliseconds} ms");

            // Save the workbook (contains the data from the last test)
            workbook.Save("PopulationPerformance.xlsx");
        }
    }
}
