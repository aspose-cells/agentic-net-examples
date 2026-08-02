// Title: Aspose.Cells .NET Benchmark: Direct Cell Formula Evaluation vs Calculation Chain
// Description: Creates a 2,000‑row workbook, fills column A with numbers, adds dependent formulas in column B, and measures execution time for three scenarios: (1) per‑cell CalculateFormula, (2) workbook.CalculateFormula with the calculation chain disabled, and (3) workbook.CalculateFormula with the chain enabled. Results are printed and the workbooks are saved for further analysis.
// Keywords: Aspose.Cells benchmark | formula calculation performance | direct cell evaluation .NET | EnableCalculationChain | disable calculation chain | CalculateFormula timing | large spreadsheet performance | C# Aspose.Cells example
// Common Searches: Aspose.Cells benchmark direct formula evaluation | how to disable calculation chain in Aspose.Cells | measure formula calculation speed Aspose.Cells .NET | compare workbook.CalculateFormula with and without chain | performance testing Aspose.Cells formulas
// Developer Intent: Compare the runtime of individual cell formula evaluation against workbook‑level calculation with the calculation chain turned on or off.
// Use Cases: Identify the fastest calculation mode for spreadsheets containing thousands of inter‑dependent formulas. | Quantify the performance impact of the EnableCalculationChain setting in real‑world workloads. | Generate reproducible timing reports for optimization or capacity‑planning purposes.
// AI Prompts: Generate a C# script that runs the benchmark 10 times, records each duration, and outputs average times for direct evaluation, no‑chain, and with‑chain calculations. | Show how to parallelize the per‑cell formula evaluation using Task Parallel Library and compare the results with the single‑threaded approach. | Explain how to interpret the benchmark output to decide when disabling the calculation chain yields the best performance.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Creates a 2,000‑row workbook, fills column A with numbers, adds dependent formulas in column B, and measures execution time for three scenarios: (1) per‑cell CalculateFormula, (2) workbook.CalculateFormula with the calculation chain disabled, and (3) workbook.CalculateFormula with the chain enabled. Results are printed and the workbooks are saved for further analysis.
class FormulaCalculationBenchmark
{
    static void Main()
    {
        try
        {
            // Create a workbook and populate it with a large set of formulas
            Workbook sourceWb = new Workbook();
            Worksheet sheet = sourceWb.Worksheets[0];
            Cells cells = sheet.Cells;

            int rowCount = 2000; // adjust for desired size

            // Initialize first column with values
            for (int i = 0; i < rowCount; i++)
            {
                cells[i, 0].PutValue(i + 1); // A column
            }

            // Add formulas that depend on the previous row (simple chain)
            for (int i = 1; i < rowCount; i++)
            {
                // B column: =A{i}+B{i-1}
                string formula = $"=A{i + 1}+B{i}";
                cells[i, 1].Formula = formula;
            }

            // Ensure the first formula cell has a base value
            cells[0, 1].Formula = "=A1*2";

            // -----------------------------------------------------------------
            // Benchmark: Direct evaluation (cell‑by‑cell)
            // -----------------------------------------------------------------
            Worksheet directSheet = sourceWb.Worksheets[0]; // use same sheet
            Stopwatch swDirect = Stopwatch.StartNew();

            for (int i = 0; i < rowCount; i++)
            {
                // Calculate each formula individually if the cell has a formula
                if (directSheet.Cells[i, 1].IsFormula)
                {
                    // CalculateFormula returns the result; we assign it back to the cell
                    object result = directSheet.CalculateFormula(directSheet.Cells[i, 1].Formula);
                    directSheet.Cells[i, 1].PutValue(result);
                }
            }

            swDirect.Stop();
            Console.WriteLine($"Direct evaluation time: {swDirect.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Helper to clone a workbook via memory stream (avoids missing Copy overload)
            // -----------------------------------------------------------------
            Workbook CloneWorkbook(Workbook original)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    original.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0;
                    return new Workbook(ms);
                }
            }

            // -----------------------------------------------------------------
            // Benchmark: Workbook calculation with calculation chain disabled
            // -----------------------------------------------------------------
            Workbook wbNoChain = CloneWorkbook(sourceWb);
            wbNoChain.Settings.FormulaSettings.EnableCalculationChain = false;

            Stopwatch swNoChain = Stopwatch.StartNew();
            wbNoChain.CalculateFormula();
            swNoChain.Stop();

            Console.WriteLine($"Workbook calculation without chain: {swNoChain.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Benchmark: Workbook calculation with calculation chain enabled
            // -----------------------------------------------------------------
            Workbook wbWithChain = CloneWorkbook(sourceWb);
            wbWithChain.Settings.FormulaSettings.EnableCalculationChain = true;

            // First calculation will build the chain; measure the total time
            Stopwatch swWithChain = Stopwatch.StartNew();
            wbWithChain.CalculateFormula();
            swWithChain.Stop();

            Console.WriteLine($"Workbook calculation with chain: {swWithChain.ElapsedMilliseconds} ms");

            // Optional: Save the workbooks (ensure the directory exists)
            string noChainPath = "Benchmark_NoChain.xlsx";
            string withChainPath = "Benchmark_WithChain.xlsx";

            wbNoChain.Save(noChainPath, SaveFormat.Xlsx);
            wbWithChain.Save(withChainPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
