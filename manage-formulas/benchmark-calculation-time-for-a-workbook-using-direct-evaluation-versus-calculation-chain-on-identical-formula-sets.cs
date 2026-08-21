// Title: Aspose.Cells .NET Benchmark: Direct Formula Evaluation vs Calculation Chain
// Description: Creates a 2,000‑row workbook where column A holds numbers and column B contains formulas that reference the previous B cell and the current A cell. The template is cloned twice: one clone evaluates each formula individually with Worksheet.CalculateFormula(string) while measuring elapsed time, and the other enables Settings.FormulaSettings.EnableCalculationChain and runs Workbook.CalculateFormula() to benchmark the full‑sheet calculation, the first run, and an incremental run after modifying A1. Both workbooks are saved for result comparison.
// Keywords: Aspose.Cells | .NET | C# | formula calculation benchmark | direct evaluation performance | calculation chain | incremental recalculation | workbook speed testing | spreadsheet processing performance | large workbook formulas
// Common Searches: Aspose.Cells benchmark formula calculation speed | direct Worksheet.CalculateFormula vs calculation chain | measure incremental recalculation time Aspose.Cells | performance test for large Excel workbooks .NET | how to enable calculation chain Aspose.Cells
// Developer Intent: Compare execution time of per‑cell formula evaluation against the calculation‑chain engine for identical formula sets.
// Use Cases: Identify the most efficient calculation method for workbooks with thousands of inter‑dependent formulas. | Evaluate the overhead of incremental updates when the calculation chain is active. | Gather performance data to guide architecture decisions for spreadsheet‑heavy services.
// AI Prompts: Rewrite the benchmark to log both elapsed milliseconds and CPU usage for each calculation method. | Interpret typical benchmark results and recommend thresholds for choosing direct evaluation versus the calculation chain in Aspose.Cells. | Extend the sample to run multiple workbook sizes and export timing results to a CSV file.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Creates a 2,000‑row workbook where column A holds numbers and column B contains formulas that reference the previous B cell and the current A cell. The template is cloned twice: one clone evaluates each formula individually with Worksheet.CalculateFormula(string) while measuring elapsed time, and the other enables Settings.FormulaSettings.EnableCalculationChain and runs Workbook.CalculateFormula() to benchmark the full‑sheet calculation, the first run, and an incremental run after modifying A1. Both workbooks are saved for result comparison.
class FormulaCalculationBenchmark
{
    static void Main()
    {
        try
        {
            // Number of rows with formulas to generate
            const int rowCount = 2000;

            // -----------------------------------------------------------------
            // Prepare a workbook with a large set of formulas (same for both tests)
            // -----------------------------------------------------------------
            Workbook wbTemplate = new Workbook();
            Worksheet wsTemplate = wbTemplate.Worksheets[0];
            Cells cells = wsTemplate.Cells;

            // Fill column A with base values
            for (int i = 0; i < rowCount; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1, A2, ...
            }

            // Add formulas in column B that depend on the previous row in column B
            // B1 = A1 * 2
            // B2 = B1 + A2
            // B3 = B2 + A3 ... etc.
            cells[0, 1].Formula = "=A1*2";
            for (int i = 1; i < rowCount; i++)
            {
                // Example: B{i+1} = B{i} + A{i+1}
                string formula = $"=B{i}+A{i + 1}";
                cells[i, 1].Formula = formula;
            }

            // -----------------------------------------------------------------
            // Benchmark: Direct evaluation (evaluate each formula individually)
            // -----------------------------------------------------------------
            // Clone the template workbook to avoid side‑effects
            Workbook wbDirect = new Workbook();
            wbDirect.Copy(wbTemplate);
            Worksheet wsDirect = wbDirect.Worksheets[0];

            Stopwatch swDirect = Stopwatch.StartNew();

            // Iterate through all formula cells and evaluate them using Worksheet.CalculateFormula(string)
            // This does not rely on the calculation chain.
            for (int i = 0; i < rowCount; i++)
            {
                string formula = wsDirect.Cells[i, 1].Formula;
                // Calculate the formula; result is returned but we don't need to store it
                wsDirect.CalculateFormula(formula);
            }

            swDirect.Stop();
            Console.WriteLine($"Direct evaluation time: {swDirect.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Benchmark: Calculation chain (enable chain and calculate whole workbook)
            // -----------------------------------------------------------------
            // Clone the template workbook again
            Workbook wbChain = new Workbook();
            wbChain.Copy(wbTemplate);
            // Enable calculation chain
            wbChain.Settings.FormulaSettings.EnableCalculationChain = true;

            Stopwatch swChain = Stopwatch.StartNew();

            // First calculation builds the chain and evaluates all formulas
            wbChain.CalculateFormula();

            swChain.Stop();
            Console.WriteLine($"Calculation chain time (first run, chain built): {swChain.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Optional: Measure subsequent calculation after a small change
            // -----------------------------------------------------------------
            // Change a single cell value to trigger incremental calculation
            wbChain.Worksheets[0].Cells[0, 0].PutValue(999); // modify A1

            Stopwatch swChainIncremental = Stopwatch.StartNew();

            // Re‑calculate; with the chain enabled only affected cells should be recomputed
            wbChain.CalculateFormula();

            swChainIncremental.Stop();
            Console.WriteLine($"Calculation chain incremental update time: {swChainIncremental.ElapsedMilliseconds} ms");

            // -----------------------------------------------------------------
            // Save workbooks (demonstrates lifecycle usage)
            // -----------------------------------------------------------------
            string directPath = "DirectEvaluationResult.xlsx";
            string chainPath = "ChainCalculationResult.xlsx";

            // Ensure the directories exist (guard for custom paths)
            string directDir = Path.GetDirectoryName(directPath);
            if (!string.IsNullOrEmpty(directDir) && !Directory.Exists(directDir))
                Directory.CreateDirectory(directDir);

            string chainDir = Path.GetDirectoryName(chainPath);
            if (!string.IsNullOrEmpty(chainDir) && !Directory.Exists(chainDir))
                Directory.CreateDirectory(chainDir);

            wbDirect.Save(directPath, SaveFormat.Xlsx);
            wbChain.Save(chainPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
