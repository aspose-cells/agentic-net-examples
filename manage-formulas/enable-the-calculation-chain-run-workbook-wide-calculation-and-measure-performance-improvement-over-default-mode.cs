// Title: Enable Calculation Chain in Aspose.Cells .NET, Run Workbook‑Wide Calculation, and Benchmark Performance
// Description: Creates a 2,000‑row workbook with dependent formulas, measures full‑workbook calculation time with the calculation chain disabled, then enables the chain, changes a single cell, runs a partial recalculation, and compares the elapsed times to demonstrate speed gains before saving the file.
// Keywords: Aspose.Cells | EnableCalculationChain | FormulaSettings | CalculateFormula | partial recalculation | performance benchmark | .NET | C# | Excel formula engine | calculation chain speed | measure calculation time
// Common Searches: how to enable calculation chain in Aspose.Cells .NET | Aspose.Cells performance test CalculateFormula | partial formula recalculation after cell update Aspose.Cells | benchmark workbook calculation with and without chain | speed up Excel formula engine using Aspose.Cells
// Developer Intent: Toggle the calculation chain, recalculate formulas, and compare execution time against the default configuration.
// Use Cases: Accelerate partial recalculations after a single‑cell edit in large workbooks. | Establish baseline and optimized timings for full‑workbook formula evaluation. | Validate that enabling the calculation chain does not affect workbook saving or lifecycle rules.
// AI Prompts: Show C# code to switch EnableCalculationChain on and off and capture CalculateFormula duration. | Explain how Aspose.Cells tracks dependencies when the calculation chain is enabled. | Generate a performance summary comparing calculation times for a workbook with 2,000 dependent rows, with the chain disabled vs. enabled.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCalculationChainDemo
{
    // Creates a 2,000‑row workbook with dependent formulas, measures full‑workbook calculation time with the calculation chain disabled, then enables the chain, changes a single cell, runs a partial recalculation, and compares the elapsed times to demonstrate speed gains before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large number of dependent formulas (e.g., 2000 rows)
            // Column A will hold base values, Column B will contain cumulative sums,
            // Column C will reference B to create a dependency chain.
            int rowCount = 2000;
            for (int i = 0; i < rowCount; i++)
            {
                // Base value in column A
                cells[i, 0].PutValue(i + 1);

                // Cumulative sum in column B: =SUM(A1:A{row})
                cells[i, 1].Formula = $"=SUM(A1:A{i + 1})";

                // Dependent formula in column C: =B{i+1}*2
                cells[i, 2].Formula = $"=B{i + 1}*2";
            }

            // -----------------------------------------------------------------
            // 1. Calculate with default settings (calculation chain disabled)
            // -----------------------------------------------------------------
            // Ensure the chain is disabled (default)
            workbook.Settings.FormulaSettings.EnableCalculationChain = false;

            // Measure calculation time
            Stopwatch sw = Stopwatch.StartNew();
            workbook.CalculateFormula(); // full workbook calculation
            sw.Stop();
            long timeWithoutChain = sw.ElapsedMilliseconds;
            Console.WriteLine($"Calculation time without chain: {timeWithoutChain} ms");

            // -----------------------------------------------------------------
            // 2. Enable calculation chain and recalculate after a small change
            // -----------------------------------------------------------------
            // Enable the calculation chain
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Modify a single cell to trigger partial recalculation
            cells[0, 0].PutValue(9999); // change A1

            // Measure recalculation time with the chain active
            sw.Restart();
            workbook.CalculateFormula(); // only affected cells should be recomputed
            sw.Stop();
            long timeWithChain = sw.ElapsedMilliseconds;
            Console.WriteLine($"Calculation time with chain after small change: {timeWithChain} ms");

            // -----------------------------------------------------------------
            // Output performance comparison
            // -----------------------------------------------------------------
            double improvement = (timeWithoutChain - timeWithChain) / (double)timeWithoutChain * 100;
            Console.WriteLine($"Performance improvement: {improvement:F2}%");

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("CalculationChainResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
