using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsPerformanceDemo
{
    class Program
    {
        static void Main()
        {
            // Parameters for the test
            const int totalRows = 1000;
            const int totalCols = 1000;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------- Row‑major population --------------------
            Stopwatch swRow = Stopwatch.StartNew();

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalCols; col++)
                {
                    // Example value: row index + column index
                    cells[row, col].PutValue(row + col);
                }
            }

            swRow.Stop();
            Console.WriteLine($"Row‑major population time: {swRow.ElapsedMilliseconds} ms");

            // Clear the sheet before the next test
            cells.Clear();

            // -------------------- Column‑major population --------------------
            Stopwatch swCol = Stopwatch.StartNew();

            for (int col = 0; col < totalCols; col++)
            {
                for (int row = 0; row < totalRows; row++)
                {
                    cells[row, col].PutValue(row + col);
                }
            }

            swCol.Stop();
            Console.WriteLine($"Column‑major population time: {swCol.ElapsedMilliseconds} ms");

            // Save the workbook (the data from the last test will be saved)
            workbook.Save("PopulationPerformance.xlsx");
        }
    }
}