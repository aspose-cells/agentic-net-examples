// Title: Measure Aspose.Cells worksheet formula calculation speed with and without the calculation chain (C#/.NET)
// Description: A C# example that builds a 2,000‑row workbook, adds dependent formulas across five columns, toggles the EnableCalculationChain setting, recalculates the worksheet using Worksheet.CalculateFormula, and records the elapsed milliseconds with Stopwatch for both enabled and disabled states.
// Keywords: Aspose.Cells calculation chain | Disable calculation chain | Worksheet.CalculateFormula performance | C# spreadsheet benchmark | formula evaluation timing .NET | Aspose.Cells performance testing
// Common Searches: Aspose.Cells benchmark calculation chain | How to measure formula calculation time in Aspose.Cells | EnableCalculationChain true vs false performance | C# speed test for spreadsheet formulas | Disable formula dependency chain Aspose.Cells
// Developer Intent: Compare the execution time of worksheet formula recalculation when the calculation chain is enabled versus when it is disabled.
// Use Cases: Determine whether turning off the calculation chain speeds up bulk updates in large workbooks. | Profile formula evaluation to choose optimal settings for automated report generation. | Benchmark Aspose.Cells performance before deploying to production environments.
// AI Prompts: Write C# code that disables the calculation chain, recalculates a specific worksheet, and logs the elapsed time using Aspose.Cells. | Explain the impact of EnableCalculationChain on formula dependency processing and when disabling it improves performance. | Provide a step‑by‑step guide to benchmark worksheet calculation time with different CalculationOptions and chain settings in Aspose.Cells.

using System;
using System.Diagnostics;
using Aspose.Cells;

// A C# example that builds a 2,000‑row workbook, adds dependent formulas across five columns, toggles the EnableCalculationChain setting, recalculates the worksheet using Worksheet.CalculateFormula, and records the elapsed milliseconds with Stopwatch for both enabled and disabled states.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate a large number of formulas for performance testing
        int totalRows = 2000;
        int totalCols = 5;

        // Fill column A with numeric values
        for (int r = 0; r < totalRows; r++)
        {
            cells[r, 0].PutValue(r + 1);
        }

        // Add formulas in column B that sum the values in column A up to the current row
        for (int r = 0; r < totalRows; r++)
        {
            cells[r, 1].Formula = $"=SUM(A1:A{r + 1})";
        }

        // Add additional simple formulas to increase workload
        for (int r = 0; r < totalRows; r++)
        {
            cells[r, 2].Formula = $"=B{r + 1}*2";
            cells[r, 3].Formula = $"=C{r + 1}+10";
            cells[r, 4].Formula = $"=D{r + 1}/3";
        }

        // Helper method to calculate the worksheet and measure elapsed time
        void MeasureCalculation(bool enableChain)
        {
            // Enable or disable the calculation chain
            workbook.Settings.FormulaSettings.EnableCalculationChain = enableChain;

            // Prepare calculation options (default options are sufficient here)
            CalculationOptions calcOptions = new CalculationOptions();

            // Measure execution time
            Stopwatch sw = Stopwatch.StartNew();

            // Recalculate all formulas in the worksheet (recursive = true)
            worksheet.CalculateFormula(calcOptions, true);

            sw.Stop();
            Console.WriteLine($"EnableCalculationChain = {enableChain}: {sw.ElapsedMilliseconds} ms");
        }

        // First measurement: calculation chain disabled
        MeasureCalculation(false);

        // Second measurement: calculation chain enabled
        MeasureCalculation(true);
    }
}
