// Title: Aspose.Cells C# Benchmark: Cell.Calculate latency with and without EnableFastFormulaCalculation across workbook sizes
// Description: A C# console program that builds three workbooks (500×500, 1000×1000, 1500×1500), fills the first row and column with constants and the remaining cells with simple sum formulas, then measures the time of a single Cell.Calculate call on the bottom‑right cell while toggling the EnableFastFormulaCalculation setting via reflection. The output shows milliseconds for each size with fast calculation ON and OFF.
// Keywords: Aspose.Cells | Cell.Calculate | EnableFastFormulaCalculation | benchmark | C# performance | formula calculation latency | large workbook | 500x500 | 1000x1000 | 1500x1500 | .NET | reflection
// Common Searches: Aspose.Cells benchmark Cell.Calculate | EnableFastFormulaCalculation performance test | measure formula calculation time C# | Cell.Calculate latency 1000x1000 workbook | how to toggle fast formula calculation Aspose.Cells | C# code to compare formula engine speed
// Developer Intent: Compare execution time of Cell.Calculate with fast formula calculation disabled versus enabled for different workbook dimensions.
// Use Cases: Identify optimal EnableFastFormulaCalculation setting for high‑volume spreadsheet processing. | Detect performance regressions after upgrading Aspose.Cells. | Profile single‑cell calculation time to guide memory and CPU budgeting. | Generate automated performance reports for CI pipelines.
// AI Prompts: Write a C# console app that creates 500×500, 1000×1000 and 1500×1500 worksheets, fills them with simple sum formulas, and logs Cell.Calculate time with EnableFastFormulaCalculation true and false using Aspose.Cells. | Explain the impact of EnableFastFormulaCalculation on Aspose.Cells’ formula engine and suggest when to use it. | Create an NUnit test that asserts Cell.Calculate with fast calculation enabled is at least 20 % faster than disabled for a 1500×1500 workbook. | Provide a PowerShell script to run the benchmark on multiple machines and aggregate results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // A C# console program that builds three workbooks (500×500, 1000×1000, 1500×1500), fills the first row and column with constants and the remaining cells with simple sum formulas, then measures the time of a single Cell.Calculate call on the bottom‑right cell while toggling the EnableFastFormulaCalculation setting via reflection. The output shows milliseconds for each size with fast calculation ON and OFF.
    class Program
    {
        // Sizes of workbooks to test (rows x columns)
        static readonly (int rows, int cols)[] WorkbookSizes = new (int, int)[]
        {
            (500, 500),
            (1000, 1000),
            (1500, 1500)
        };

        static void Main()
        {
            // Iterate over each workbook size
            foreach (var size in WorkbookSizes)
            {
                int rows = size.rows;
                int cols = size.cols;

                // Create workbook and fill with simple formulas
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Populate cells with a formula that sums the cell to the left and above
                // First row and column get constant values
                for (int i = 0; i < rows; i++)
                {
                    cells[i, 0].PutValue(1); // first column constant
                }
                for (int j = 0; j < cols; j++)
                {
                    cells[0, j].PutValue(1); // first row constant
                }

                // Fill the rest with formulas
                for (int i = 1; i < rows; i++)
                {
                    for (int j = 1; j < cols; j++)
                    {
                        cells[i, j].Formula = $"=A{i + 1}+B{j + 1}";
                    }
                }

                // Choose a target cell for individual calculation
                Cell targetCell = cells[rows - 1, cols - 1];

                // Benchmark without fast formula calculation
                wb.Settings.FormulaSettings.EnableCalculationChain = false; // ensure chain disabled
                // Assume the existence of EnableFastFormulaCalculation property; set to false if present
                SetEnableFastFormulaCalculation(wb, false);
                RunCellCalculateBenchmark(wb, targetCell, "FastCalc OFF");

                // Benchmark with fast formula calculation enabled
                wb.Settings.FormulaSettings.EnableCalculationChain = false;
                SetEnableFastFormulaCalculation(wb, true);
                RunCellCalculateBenchmark(wb, targetCell, "FastCalc ON");
            }
        }

        // Helper to set EnableFastFormulaCalculation if the property exists
        static void SetEnableFastFormulaCalculation(Workbook wb, bool enable)
        {
            // The property may not exist in older versions; use reflection to avoid compile errors
            var formulaSettings = wb.Settings.FormulaSettings;
            var prop = formulaSettings.GetType().GetProperty("EnableFastFormulaCalculation");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(formulaSettings, enable);
            }
        }

        // Executes the Cell.Calculate measurement and prints elapsed time
        static void RunCellCalculateBenchmark(Workbook wb, Cell cell, string label)
        {
            // Ensure any previous calculations are cleared
            wb.CalculateFormula(); // warm‑up to build any internal structures

            // Measure the time for a single Cell.Calculate call
            Stopwatch sw = Stopwatch.StartNew();
            cell.Calculate(new CalculationOptions());
            sw.Stop();

            Console.WriteLine($"{label}: Workbook {wb.Worksheets[0].Cells.MaxDataRow + 1}x{wb.Worksheets[0].Cells.MaxDataColumn + 1} - Cell.Calculate took {sw.ElapsedMilliseconds} ms");
        }
    }
}
