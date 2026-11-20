// Title: Benchmark column‑major vs row‑major cell population speed with Aspose.Cells for .NET
// Common Searches: how to benchmark column-major cell filling versus row-major in Aspose.Cells C# | measure performance of populating a 1000x1000 worksheet with Aspose.Cells | Aspose.Cells best practice for fast data insertion into large Excel sheets | using Stopwatch to compare column-wise and row-wise write speed in Aspose.Cells | optimal cell population order for Aspose.Cells performance testing
// Tags: column-major cell insertion performance Aspose.Cells | row-major worksheet data population timing Aspose.Cells | Aspose.Cells large worksheet benchmarking C# | Stopwatch execution time measurement Aspose.Cells | Excel data write speed optimization Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceTest
{
    // The example creates a 1000 × 1000 worksheet, times column‑by‑column and row‑by‑row cell population using Stopwatch, clears the sheet between runs, prints elapsed milliseconds for each method, and saves the workbook to verify the data.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                const int rows = 1000; // number of rows
                const int cols = 1000; // number of columns

                // ------------------------------
                // Column‑major population (fill column by column)
                // ------------------------------
                Stopwatch swColumn = Stopwatch.StartNew();

                for (int col = 0; col < cols; col++)
                {
                    for (int row = 0; row < rows; row++)
                    {
                        // Example value: row index + column index
                        cells[row, col].PutValue(row + col);
                    }
                }

                swColumn.Stop();
                Console.WriteLine($"Column‑major population time: {swColumn.ElapsedMilliseconds} ms");

                // Clear the sheet for the next test
                sheet.Cells.Clear(); // clears all cells (values, formulas, formatting)

                // ------------------------------
                // Row‑major population (fill row by row)
                // ------------------------------
                Stopwatch swRow = Stopwatch.StartNew();

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        cells[row, col].PutValue(row + col);
                    }
                }

                swRow.Stop();
                Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

                // Save the workbook (optional, just to verify that data was written)
                workbook.Save("PopulationPerformance.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
