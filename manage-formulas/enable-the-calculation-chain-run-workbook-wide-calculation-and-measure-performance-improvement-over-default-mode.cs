// Title: Aspose.Cells .NET – Benchmark Calculation Chain vs No Chain for Workbook‑wide Formula Evaluation
// Description: This C# example builds two identical workbooks with 1,000 rows of inter‑dependent formulas, runs workbook‑wide CalculateFormula once with the calculation chain disabled and once with it enabled, measures execution time using Stopwatch, and outputs the results to illustrate the performance benefit of EnableCalculationChain.
// Keywords: Aspose.Cells | .NET | C# | EnableCalculationChain | CalculationChainPerformanceDemo | CalculateFormula | formula performance benchmark | spreadsheet calculation speed | large dependent formulas | Excel performance tuning
// Common Searches: Aspose.Cells calculation chain performance | EnableCalculationChain benchmark C# | CalculateFormula speed test Aspose.Cells | measure formula calculation time Aspose.Cells | impact of calculation chain on large spreadsheets
// Developer Intent: Measure and compare the execution time of workbook‑wide formula calculation with the calculation chain disabled versus enabled in Aspose.Cells for .NET.
// Use Cases: Benchmarking the effect of the calculation chain on spreadsheets with deep formula dependencies. | Deciding whether to enable the calculation chain to improve performance of large Excel files. | Validating that enabling the calculation chain yields identical results faster than the default mode.
// AI Prompts: Write C# code that creates a workbook with 10,000 dependent formulas, runs CalculateFormula with EnableCalculationChain set to false and true, and logs the elapsed milliseconds for each run. | Explain how Aspose.Cells' calculation chain works and why it can accelerate evaluation of long formula dependency chains. | Provide a step‑by‑step tutorial for measuring and logging formula calculation performance when toggling EnableCalculationChain in Aspose.Cells settings.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example builds two identical workbooks with 1,000 rows of inter‑dependent formulas, runs workbook‑wide CalculateFormula once with the calculation chain disabled and once with it enabled, measures execution time using Stopwatch, and outputs the results to illustrate the performance benefit of EnableCalculationChain.
    class CalculationChainPerformanceDemo
    {
        static void Main()
        {
            // Create a workbook and fill it with a large number of dependent formulas
            Workbook wbNoChain = new Workbook();
            Worksheet wsNoChain = wbNoChain.Worksheets[0];
            Cells cellsNoChain = wsNoChain.Cells;

            // Populate column A with base values
            for (int i = 0; i < 1000; i++)
            {
                cellsNoChain[i, 0].PutValue(i + 1); // A1..A1000
            }

            // Create dependent formulas in column B that sum a range in column A
            // Each B cell depends on the previous B cell, forming a long dependency chain
            cellsNoChain[0, 1].Formula = "=SUM(A1:A1)"; // B1
            for (int i = 1; i < 1000; i++)
            {
                // B(i+1) = B(i) + A(i+1)
                cellsNoChain[i, 1].Formula = $"=B{i}+A{i + 1}";
            }

            // ------------------------------
            // 1. Calculate without calculation chain (default)
            // ------------------------------
            wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false; // explicit for clarity
            Stopwatch swNoChain = Stopwatch.StartNew();
            wbNoChain.CalculateFormula(); // workbook‑wide calculation
            swNoChain.Stop();

            // ------------------------------
            // 2. Calculate with calculation chain enabled
            // ------------------------------
            // Create a fresh workbook with the same data to avoid cached results
            Workbook wbWithChain = new Workbook();
            Worksheet wsWithChain = wbWithChain.Worksheets[0];
            Cells cellsWithChain = wsWithChain.Cells;

            // Copy the same data and formulas
            for (int i = 0; i < 1000; i++)
            {
                cellsWithChain[i, 0].PutValue(i + 1);
            }
            cellsWithChain[0, 1].Formula = "=SUM(A1:A1)";
            for (int i = 1; i < 1000; i++)
            {
                cellsWithChain[i, 1].Formula = $"=B{i}+A{i + 1}";
            }

            // Enable calculation chain
            wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true;
            Stopwatch swWithChain = Stopwatch.StartNew();
            wbWithChain.CalculateFormula(); // first calculation builds the chain
            swWithChain.Stop();

            // Output the measured times
            Console.WriteLine($"Calculation time without chain: {swNoChain.ElapsedMilliseconds} ms");
            Console.WriteLine($"Calculation time with chain   : {swWithChain.ElapsedMilliseconds} ms");

            // Optional: save the workbooks to verify results (uses the allowed save lifecycle)
            wbNoChain.Save("NoChainResult.xlsx", SaveFormat.Xlsx);
            wbWithChain.Save("WithChainResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
