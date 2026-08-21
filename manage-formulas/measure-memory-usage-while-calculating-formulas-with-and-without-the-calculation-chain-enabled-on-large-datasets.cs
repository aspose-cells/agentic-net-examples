// Title: Measure Aspose.Cells Formula Calculation Memory Usage With and Without Calculation Chain (Large Workbook)
// Description: C# example that creates two 5,000‑row × 10‑column workbooks, fills them with values and dependent formulas, toggles the EnableCalculationChain setting, forces garbage collection, records memory before and after CalculateFormula, measures elapsed time with Stopwatch, prints the results and saves the files. Use it to benchmark memory and performance impact of the calculation chain on large datasets.
// Keywords: Aspose.Cells | C# | memory usage | formula calculation | calculation chain | EnableCalculationChain | benchmark | large workbook | performance | GC.Collect | Stopwatch | Excel | XLSX
// Common Searches: Aspose.Cells memory usage formula calculation | EnableCalculationChain performance impact | measure memory before and after CalculateFormula | benchmark Aspose.Cells formula engine | how to disable calculation chain in Aspose.Cells | memory consumption large workbook Aspose.Cells
// Developer Intent: Evaluate how enabling or disabling the calculation chain affects memory consumption and execution time when calculating formulas in a large Aspose.Cells workbook.
// Use Cases: Determine baseline memory usage when the calculation chain is turned off. | Quantify the additional memory and time overhead introduced by the calculation chain. | Guide configuration decisions for performance‑critical applications that process large Excel files with Aspose.Cells.
// AI Prompts: Generate C# code that logs memory delta and elapsed time for Aspose.Cells.CalculateFormula with EnableCalculationChain toggled on a 10,000‑row workbook. | Explain how the calculation chain influences memory allocation and CPU usage during formula evaluation in Aspose.Cells. | Create a script to run multiple iterations with varying workbook sizes, collect memory and timing data, and output a CSV for plotting the impact of EnableCalculationChain.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    // C# example that creates two 5,000‑row × 10‑column workbooks, fills them with values and dependent formulas, toggles the EnableCalculationChain setting, forces garbage collection, records memory before and after CalculateFormula, measures elapsed time with Stopwatch, prints the results and saves the files. Use it to benchmark memory and performance impact of the calculation chain on large datasets.
    class Program
    {
        static void Main()
        {
            // Parameters for the large dataset
            const int rows = 5000;
            const int cols = 10;

            // ------------------------------------------------------------
            // Scenario 1: Calculation chain disabled
            // ------------------------------------------------------------
            Workbook wbNoChain = new Workbook();
            Worksheet wsNoChain = wbNoChain.Worksheets[0];
            Cells cellsNoChain = wsNoChain.Cells;

            // Populate cells with values and dependent formulas
            for (int r = 0; r < rows; r++)
            {
                // First column gets a simple value
                cellsNoChain[r, 0].PutValue(r + 1);

                // Remaining columns contain formulas that depend on the previous column
                for (int c = 1; c < cols; c++)
                {
                    // Example: =A1*2, =B1*2, etc.
                    string prevColLetter = CellIndexToName(r, c - 1).Substring(0, 1);
                    string formula = $"={prevColLetter}{r + 1}*2";
                    cellsNoChain[r, c].Formula = formula;
                }
            }

            // Disable calculation chain
            wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false;

            // Ensure a clean memory baseline
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memBeforeNoChain = GC.GetTotalMemory(true);
            Stopwatch swNoChain = Stopwatch.StartNew();

            // Perform calculation
            wbNoChain.CalculateFormula();

            swNoChain.Stop();
            long memAfterNoChain = GC.GetTotalMemory(true);

            Console.WriteLine("=== Calculation Chain Disabled ===");
            Console.WriteLine($"Time elapsed: {swNoChain.ElapsedMilliseconds} ms");
            Console.WriteLine($"Memory before: {memBeforeNoChain / 1024} KB");
            Console.WriteLine($"Memory after : {memAfterNoChain / 1024} KB");
            Console.WriteLine($"Memory increase: {(memAfterNoChain - memBeforeNoChain) / 1024} KB");
            Console.WriteLine();

            // Save the workbook (uses the provided save rule)
            wbNoChain.Save("LargeDataset_NoChain.xlsx", SaveFormat.Xlsx);

            // ------------------------------------------------------------
            // Scenario 2: Calculation chain enabled
            // ------------------------------------------------------------
            Workbook wbWithChain = new Workbook();
            Worksheet wsWithChain = wbWithChain.Worksheets[0];
            Cells cellsWithChain = wsWithChain.Cells;

            // Populate the same data pattern
            for (int r = 0; r < rows; r++)
            {
                cellsWithChain[r, 0].PutValue(r + 1);
                for (int c = 1; c < cols; c++)
                {
                    string prevColLetter = CellIndexToName(r, c - 1).Substring(0, 1);
                    string formula = $"={prevColLetter}{r + 1}*2";
                    cellsWithChain[r, c].Formula = formula;
                }
            }

            // Enable calculation chain
            wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true;

            // Clean memory again
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long memBeforeWithChain = GC.GetTotalMemory(true);
            Stopwatch swWithChain = Stopwatch.StartNew();

            // Perform calculation (first calculation will also build the chain)
            wbWithChain.CalculateFormula();

            swWithChain.Stop();
            long memAfterWithChain = GC.GetTotalMemory(true);

            Console.WriteLine("=== Calculation Chain Enabled ===");
            Console.WriteLine($"Time elapsed: {swWithChain.ElapsedMilliseconds} ms");
            Console.WriteLine($"Memory before: {memBeforeWithChain / 1024} KB");
            Console.WriteLine($"Memory after : {memAfterWithChain / 1024} KB");
            Console.WriteLine($"Memory increase: {(memAfterWithChain - memBeforeWithChain) / 1024} KB");
            Console.WriteLine();

            // Save the workbook
            wbWithChain.Save("LargeDataset_WithChain.xlsx", SaveFormat.Xlsx);
        }

        // Helper method to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 => A1)
        private static string CellIndexToName(int row, int column)
        {
            // Convert column index to letters
            string colName = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }

            // Row numbers are 1‑based in Excel
            return $"{colName}{row + 1}";
        }
    }
}
