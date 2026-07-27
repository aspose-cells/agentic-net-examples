using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Define size of the data set
            const int rows = 5000;
            const int cols = 200;

            // Prepare a 2‑dimensional array with sample data
            double[,] data = new double[rows, cols];
            Random rnd = new Random();
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    data[r, c] = rnd.NextDouble() * 1000;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // Row‑major population (row loop outer, column loop inner)
            // -----------------------------------------------------------------
            Stopwatch swRow = Stopwatch.StartNew();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }
            swRow.Stop();
            Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

            // Clear the sheet before the next test
            cells.Clear();

            // -----------------------------------------------------------------
            // Column‑major population (column loop outer, row loop inner)
            // -----------------------------------------------------------------
            Stopwatch swCol = Stopwatch.StartNew();
            for (int c = 0; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    cells[r, c].PutValue(data[r, c]);
                }
            }
            swCol.Stop();
            Console.WriteLine($"Column‑major population time: {swCol.ElapsedMilliseconds} ms");

            // Save the workbook (optional – shows that the data was written)
            workbook.Save("PopulationPerformanceDemo.xlsx");
        }
    }
}