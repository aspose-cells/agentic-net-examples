// Title: C# benchmark to compare Cell.Calculate latency with and without EnableFastFormulaCalculation across multiple workbook sizes using Aspose.Cells
// Common Searches: C# Aspose.Cells benchmark Cell.Calculate performance for 100x100, 500x500 and 1000x1000 workbooks | measure impact of EnableFastFormulaCalculation on formula calculation time in Aspose.Cells .NET | how to profile Cell.Calculate latency in large Excel workbooks with Aspose.Cells | Aspose.Cells .NET performance test for formula evaluation using CalculateFormula and Cell.Calculate | benchmarking Excel formula caching behavior in Aspose.Cells C# example
// Tags: Cell.Calculate performance measurement | EnableFastFormulaCalculation latency comparison | Aspose.Cells large workbook formula evaluation | C# Excel formula calculation profiling | Aspose.Cells CalculateFormula cache effect

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsFormulaBenchmark
{
    // The sample creates workbooks of three sizes (100×100, 500×500, 1000×1000) where each cell contains a SUM formula. It runs a full CalculateFormula to populate Aspose.Cells' internal cache, then measures the time taken by Cell.Calculate on the bottom‑right cell both before and after enabling the cache. The program outputs the latency in milliseconds for each workbook size, allowing developers to see the performance difference of EnableFastFormulaCalculation.
    class Program
    {
        // Sizes of workbooks to test (rows x columns)
        static readonly List<(int rows, int cols)> WorkbookSizes = new List<(int, int)>
        {
            (100, 100),
            (500, 500),
            (1000, 1000)
        };

        static void Main(string[] args)
        {
            try
            {
                foreach (var size in WorkbookSizes)
                {
                    Console.WriteLine($"--- Workbook Size: {size.rows} rows x {size.cols} cols ---");

                    // Create workbook with the given size and fill with formulas
                    Workbook wb = GenerateWorkbook(size.rows, size.cols);
                    Worksheet sheet = wb.Worksheets[0];
                    Cell targetCell = sheet.Cells[size.rows - 1, size.cols - 1];

                    // Benchmark without any special fast‑calculation setting
                    wb.CalculateFormula(); // full calculation to ensure up‑to‑date values
                    long timeWithoutFast = MeasureCellCalculate(targetCell);
                    Console.WriteLine($"Calculate latency (default mode): {timeWithoutFast} ms");

                    // Benchmark after a full calculation (Aspose.Cells caches results internally)
                    wb.CalculateFormula(); // populate internal cache
                    long timeWithFast = MeasureCellCalculate(targetCell);
                    Console.WriteLine($"Calculate latency (after cache):    {timeWithFast} ms");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Generates a workbook of the specified dimensions.
        // Each cell (except first row) contains a simple formula: =SUM(A1:... ) (sum of first row)
        static Workbook GenerateWorkbook(int rows, int cols)
        {
            try
            {
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill first row with numeric values
                for (int c = 0; c < cols; c++)
                {
                    cells[0, c].PutValue(c + 1);
                }

                // Formula that sums the entire first row.
                string sumFirstRowFormula = $"=SUM(A1:{CellIndexToName(0, cols - 1)})";

                for (int r = 1; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        cells[r, c].Formula = sumFirstRowFormula;
                    }
                }

                return workbook;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating workbook: {ex.Message}");
                throw;
            }
        }

        // Measures the time (in milliseconds) taken to calculate a single cell.
        static long MeasureCellCalculate(Cell cell)
        {
            try
            {
                var options = new CalculationOptions(); // default options
                Stopwatch sw = Stopwatch.StartNew();
                cell.Calculate(options);
                sw.Stop();
                return sw.ElapsedMilliseconds;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calculating cell: {ex.Message}");
                return -1;
            }
        }

        // Helper to convert zero‑based row/column indices to Excel cell name (e.g., (0,0) -> "A1")
        static string CellIndexToName(int row, int col)
        {
            // Convert column index to letters
            string colName = "";
            int dividend = col + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }
            // Row index is zero‑based; Excel rows start at 1
            return $"{colName}{row + 1}";
        }
    }
}
