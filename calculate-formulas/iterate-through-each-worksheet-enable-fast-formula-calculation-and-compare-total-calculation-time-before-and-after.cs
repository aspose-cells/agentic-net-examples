// Title: C# Benchmark: Aspose.Cells formula calculation speed with and without EnableCalculationChain across multiple worksheets
// Description: A C# sample that builds a workbook with three sheets, fills cells with numeric values and SUM/multiplication formulas, records the time taken by Workbook.CalculateFormula() before and after activating workbook.Settings.FormulaSettings.EnableCalculationChain, reports the time saved, and writes the workbook to an XLSX file.
// Keywords: Aspose.Cells C# | EnableCalculationChain | formula calculation performance | Workbook.CalculateFormula benchmark | fast formula evaluation .NET | multiple worksheets Excel | measure calculation time | Excel formula speed | Aspose.Cells performance tuning | formula dependency chain
// Common Searches: Aspose.Cells enable calculation chain performance | C# benchmark formula calculation time Aspose.Cells | How to speed up Workbook.CalculateFormula | Compare Aspose.Cells calculation speed with and without chain | Measure impact of EnableCalculationChain on large workbooks
// Developer Intent: Evaluate the performance difference of formula calculation before and after turning on the fast calculation chain in Aspose.Cells.
// Use Cases: Determine whether EnableCalculationChain provides measurable speed gains for large financial models. | Optimize automated report generation pipelines that rely on intensive Excel formula processing. | Validate performance improvements when scaling workbooks with many inter‑sheet formulas.
// AI Prompts: Generate C# code that creates a workbook with 5 worksheets, each containing 1,000 rows of SUM formulas, then measures calculation time with EnableCalculationChain disabled and enabled. | Explain how Aspose.Cells' calculation chain reduces formula recomputation overhead and improves dependency resolution. | Suggest best practices for using EnableCalculationChain in production, including when to disable it for debugging or when working with volatile functions.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaPerformanceDemo
{
    // A C# sample that builds a workbook with three sheets, fills cells with numeric values and SUM/multiplication formulas, records the time taken by Workbook.CalculateFormula() before and after activating workbook.Settings.FormulaSettings.EnableCalculationChain, reports the time saved, and writes the workbook to an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains one default worksheet)
                Workbook workbook = new Workbook();

                const int sheetCount = 3;
                const int rows = 500;
                const int cols = 20;

                // Ensure the workbook has the required number of worksheets
                for (int s = 0; s < sheetCount; s++)
                {
                    Worksheet sheet;
                    if (s < workbook.Worksheets.Count)
                    {
                        // Existing worksheet
                        sheet = workbook.Worksheets[s];
                    }
                    else
                    {
                        // Add a new worksheet (Worksheets.Add returns the index of the new sheet)
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    sheet.Name = $"Sheet{s + 1}";

                    // Fill first column with values
                    for (int r = 0; r < rows; r++)
                    {
                        sheet.Cells[r, 0].PutValue(r + 1);
                    }

                    // Create formulas that sum a range in the first column
                    for (int r = 0; r < rows; r++)
                    {
                        // Example: =SUM($A$1:A{r+1})
                        string formula = $"=SUM($A$1:A{r + 1})";
                        sheet.Cells[r, 1].Formula = formula;
                    }

                    // Add additional formulas across the row to increase complexity
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 2; c < cols; c++)
                        {
                            // Example: =B{r+1}*C{r+1}
                            string formula = $"=B{r + 1}*C{r + 1}";
                            sheet.Cells[r, c].Formula = formula;
                        }
                    }
                }

                // -----------------------------------------------------------------
                // First calculation: without fast calculation chain (default)
                // -----------------------------------------------------------------
                Stopwatch sw = new Stopwatch();
                sw.Start();

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                sw.Stop();
                long timeWithoutChain = sw.ElapsedMilliseconds;
                Console.WriteLine($"Calculation time without calculation chain: {timeWithoutChain} ms");

                // -----------------------------------------------------------------
                // Enable fast formula calculation (calculation chain) globally
                // -----------------------------------------------------------------
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;

                // -----------------------------------------------------------------
                // Second calculation: with fast calculation chain enabled
                // -----------------------------------------------------------------
                sw.Restart();

                // Re‑calculate all formulas after enabling the chain
                workbook.CalculateFormula();

                sw.Stop();
                long timeWithChain = sw.ElapsedMilliseconds;
                Console.WriteLine($"Calculation time with calculation chain: {timeWithChain} ms");

                // -----------------------------------------------------------------
                // Output comparison
                // -----------------------------------------------------------------
                Console.WriteLine($"Time saved: {timeWithoutChain - timeWithChain} ms");

                // Save the workbook
                string outputPath = "FormulaPerformanceResult.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
