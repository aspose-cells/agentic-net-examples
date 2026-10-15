// Title: Measure column‑major vs row‑major worksheet population speed with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that uses Aspose.Cells to fill a worksheet column‑first, then clears the sheet and fills it row‑first, measuring each pass with Stopwatch and printing the elapsed milliseconds. | Modify the benchmark to run each fill order multiple times, compute the average elapsed time for column‑major and row‑major loops, and display which ordering is faster. | Add command‑line arguments that let users specify the number of rows, columns, and iteration count for the population test, and log the timing results to a CSV file.
// Common Searches: how fast is column‑first cell insertion compared to row‑first using Aspose.Cells .NET | benchmark Excel worksheet population order with Aspose.Cells C# | measure write performance of large Excel sheets in .NET with Stopwatch | optimal loop order for filling massive worksheets in Aspose.Cells | Aspose.Cells performance test for 5 million cells column vs row
// Tags: Aspose.Cells column-major fill benchmark | Aspose.Cells row-major population performance | Excel worksheet write speed measurement .NET | cell write performance profiling Aspose | large worksheet data insertion optimization

using System;
using System.Diagnostics;
using Aspose.Cells;

// // Benchmarks the time required to populate a 5000x2000 worksheet using column‑major and row‑major loops with Aspose.Cells, clears the sheet between runs, outputs elapsed milliseconds for each method, and saves the workbook.
class ColumnRowPopulationBenchmark
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define dimensions for the test data
        const int rows = 5000;
        const int cols = 2000;

        // -------------------- Column‑major population --------------------
        Stopwatch swColumn = Stopwatch.StartNew();

        for (int col = 0; col < cols; col++)
        {
            for (int row = 0; row < rows; row++)
            {
                // Populate each cell with a simple value (e.g., row index)
                sheet.Cells[row, col].PutValue(row);
            }
        }

        swColumn.Stop();
        Console.WriteLine($"Column‑major population time: {swColumn.ElapsedMilliseconds} ms");

        // Clear the worksheet for the next test
        sheet.Cells.Clear();

        // -------------------- Row‑major population --------------------
        Stopwatch swRow = Stopwatch.StartNew();

        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < cols; col++)
            {
                // Populate each cell with a simple value (e.g., column index)
                sheet.Cells[row, col].PutValue(col);
            }
        }

        swRow.Stop();
        Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

        // Save the workbook (optional, just to demonstrate lifecycle compliance)
        workbook.Save("PopulationBenchmark.xlsx");
    }
}
