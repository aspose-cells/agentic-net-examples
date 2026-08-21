// Title: Benchmark Worksheet.CalculateFormula with Calculation Chain Disabled vs Enabled in Aspose.Cells for .NET (C#)
// Description: C# sample that builds a workbook with thousands of formulas, toggles Workbook.Settings.FormulaSettings.EnableCalculationChain, recalculates only the first worksheet using Worksheet.CalculateFormula, measures execution time with Stopwatch, and saves the results. Demonstrates performance impact of the calculation chain on large sheets.
// Keywords: Aspose.Cells | C# | .NET | calculation chain | EnableCalculationChain | Worksheet.CalculateFormula | formula performance | benchmark Excel calculation | recalculate single worksheet | large workbook optimization
// Common Searches: Aspose.Cells disable calculation chain performance | measure formula calculation time Aspose.Cells C# | Worksheet.CalculateFormula benchmark | how to turn off calculation chain in Aspose.Cells | speed up Excel calculations with Aspose.Cells
// Developer Intent: Evaluate the speed difference of recalculating a worksheet when the calculation chain is turned off versus on.
// Use Cases: Determine whether disabling the calculation chain improves performance for workbooks with many inter‑sheet formulas. | Recalculate only a specific sheet after data changes without triggering full workbook evaluation. | Create automated performance tests for formula engines in .NET applications.
// AI Prompts: Show how to disable the calculation chain in Aspose.Cells, recalculate a single worksheet, and log the elapsed milliseconds for both states. | Provide a C# code snippet that benchmarks Worksheet.CalculateFormula with EnableCalculationChain set to true and false. | Explain how to interpret timing results when comparing calculation chain settings in Aspose.Cells for .NET.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsCalculationChainDemo
{
    // C# sample that builds a workbook with thousands of formulas, toggles Workbook.Settings.FormulaSettings.EnableCalculationChain, recalculates only the first worksheet using Worksheet.CalculateFormula, measures execution time with Stopwatch, and saves the results. Demonstrates performance impact of the calculation chain on large sheets.
    class Program
    {
        static void Main()
        {
            try
            {
                const int rowCount = 2000;

                // Prepare a workbook with many formulas to see performance difference
                Workbook wbTemplate = CreateWorkbookWithFormulas(rowCount);

                // ----------- Test with Calculation Chain Disabled -----------
                // Create a fresh workbook with the same formulas
                Workbook wbNoChain = CreateWorkbookWithFormulas(rowCount);
                wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false;

                Stopwatch swNoChain = Stopwatch.StartNew();
                // Recalculate only the first worksheet
                Worksheet wsNoChain = wbNoChain.Worksheets[0];
                wsNoChain.CalculateFormula(new CalculationOptions(), true);
                swNoChain.Stop();

                Console.WriteLine($"Calculation time with chain disabled: {swNoChain.ElapsedMilliseconds} ms");

                // ----------- Test with Calculation Chain Enabled -----------
                // Create another fresh workbook with the same formulas
                Workbook wbWithChain = CreateWorkbookWithFormulas(rowCount);
                wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true;

                Stopwatch swWithChain = Stopwatch.StartNew();
                Worksheet wsWithChain = wbWithChain.Worksheets[0];
                wsWithChain.CalculateFormula(new CalculationOptions(), true);
                swWithChain.Stop();

                Console.WriteLine($"Calculation time with chain enabled: {swWithChain.ElapsedMilliseconds} ms");

                // Save the workbooks (optional, demonstrates lifecycle usage)
                wbNoChain.Save("NoChainResult.xlsx", SaveFormat.Xlsx);
                wbWithChain.Save("WithChainResult.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to create a workbook filled with formulas
        private static Workbook CreateWorkbookWithFormulas(int rowCount)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Put initial value in A1
            cells["A1"].PutValue(1);

            // Each subsequent cell in column A adds 1 to the previous cell
            for (int i = 2; i <= rowCount; i++)
            {
                string prevCell = $"A{i - 1}";
                string curCell = $"A{i}";
                cells[curCell].Formula = $"={prevCell}+1";
            }

            // Column B sums the range A1:A{rowCount}
            cells[$"B1"].Formula = $"=SUM(A1:A{rowCount})";

            return wb;
        }
    }
}
