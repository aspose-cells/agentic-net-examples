using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    class Program
    {
        // Number of rows and columns for the test sheet
        const int RowCount = 2000;
        const int ColCount = 10;

        static void Main()
        {
            // Benchmark without calculation chain (direct evaluation)
            Workbook wbDirect = CreateWorkbookWithFormulas();
            wbDirect.Settings.FormulaSettings.EnableCalculationChain = false;
            TimeSpan directTime = MeasureCalculationTime(wbDirect);
            Console.WriteLine($"Direct evaluation time (chain disabled): {directTime.TotalMilliseconds} ms");

            // Benchmark with calculation chain enabled
            Workbook wbChain = CreateWorkbookWithFormulas();
            wbChain.Settings.FormulaSettings.EnableCalculationChain = true;
            TimeSpan chainTime = MeasureCalculationTime(wbChain);
            Console.WriteLine($"Calculation chain time (chain enabled): {chainTime.TotalMilliseconds} ms");

            // Save workbooks (optional, demonstrates lifecycle usage)
            wbDirect.Save("Benchmark_Direct.xlsx", SaveFormat.Xlsx);
            wbChain.Save("Benchmark_Chain.xlsx", SaveFormat.Xlsx);
        }

        // Creates a new workbook and fills it with a set of formulas
        static Workbook CreateWorkbookWithFormulas()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Fill first column with base values
            for (int r = 0; r < RowCount; r++)
            {
                cells[r, 0].PutValue(r + 1); // Simple numeric values
            }

            // Generate formulas that depend on the previous column
            // Example: each cell in column c (c>0) = SUM of the same row from column 0 to c-1
            for (int r = 0; r < RowCount; r++)
            {
                for (int c = 1; c < ColCount; c++)
                {
                    string startCell = CellsHelper.CellIndexToName(r, 0);
                    string endCell = CellsHelper.CellIndexToName(r, c - 1);
                    cells[r, c].Formula = $"=SUM({startCell}:{endCell})";
                }
            }

            return wb;
        }

        // Measures the time taken to calculate all formulas in the given workbook
        static TimeSpan MeasureCalculationTime(Workbook wb)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            // Calculate all formulas in the workbook
            wb.CalculateFormula();

            sw.Stop();
            return sw.Elapsed;
        }
    }
}