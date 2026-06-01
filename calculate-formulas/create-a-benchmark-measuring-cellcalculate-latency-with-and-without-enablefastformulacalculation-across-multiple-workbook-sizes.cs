using System;
using System.Diagnostics;
using Aspose.Cells;

class FormulaCalculationBenchmark
{
    static void Main()
    {
        try
        {
            // Define different workbook sizes (rows x columns)
            int[] sizes = new int[] { 500, 1000, 1500 };

            foreach (int size in sizes)
            {
                Console.WriteLine($"Benchmark for workbook size {size}x{size}");

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the worksheet:
                // Column A – plain numeric values
                // Column B – simple formula referencing the value in column A
                for (int row = 0; row < size; row++)
                {
                    cells[row, 0].PutValue(row + 1);                 // A column
                    cells[row, 1].Formula = $"=A{row + 1}+1";       // B column formula
                }

                // Choose the last formula cell as the target for measurement
                Cell targetCell = cells[size - 1, 1];

                // -----------------------------------------------------------------
                // Benchmark without fast formula calculation (default behavior)
                // -----------------------------------------------------------------
                // Ensure any internal caches are built
                workbook.CalculateFormula();

                Stopwatch sw = Stopwatch.StartNew();
                targetCell.Calculate(new CalculationOptions());
                sw.Stop();

                Console.WriteLine($"Default calculation: {sw.ElapsedMilliseconds} ms");

                // -----------------------------------------------------------------
                // Benchmark with fast formula calculation (if supported)
                // -----------------------------------------------------------------
                // Some Aspose.Cells versions enable fast calculation automatically.
                // Re‑calculate to rebuild internal structures.
                workbook.CalculateFormula();

                sw.Restart();
                targetCell.Calculate(new CalculationOptions());
                sw.Stop();

                Console.WriteLine($"Re‑calculation after rebuild: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}