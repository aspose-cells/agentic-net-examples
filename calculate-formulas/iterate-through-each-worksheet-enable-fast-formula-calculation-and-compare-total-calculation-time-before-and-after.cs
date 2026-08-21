// Title: Benchmark Fast Formula Calculation with EnableCalculationChain across Multiple Worksheets in Aspose.Cells for .NET
// Description: This C# example builds a workbook with three worksheets, each containing 2,000 rows of numeric data, simple multiplication formulas, and cumulative SUM formulas. It measures the total calculation time using the default engine, then enables the fast formula calculation feature (EnableCalculationChain), recalculates, and records the new timing before saving the file. The sample demonstrates how to profile performance gains when processing large workbooks with Aspose.Cells.
// Keywords: Aspose.Cells fast formula calculation | EnableCalculationChain .NET | benchmark formula performance | measure workbook calculation time | Aspose.Cells C# performance test | formula calculation chain | large workbook profiling
// Common Searches: Aspose.Cells enable calculation chain performance | how to benchmark formula calculation in Aspose.Cells | measure calculation time before and after EnableCalculationChain | C# Aspose.Cells fast formula evaluation example | compare workbook calculation speed Aspose.Cells
// Developer Intent: Assess the speed improvement obtained by turning on EnableCalculationChain when evaluating formulas across all worksheets in a large Aspose.Cells workbook.
// Use Cases: Profile a heavy workbook to decide if the calculation chain should be enabled in production. | Validate that cumulative SUM formulas run faster with the fast calculation mode. | Generate timing reports for formula evaluation before and after enabling the calculation chain.
// AI Prompts: Write C# code that iterates through each worksheet, enables EnableCalculationChain, and logs calculation time per worksheet using Aspose.Cells. | Explain how to interpret the timing results and set performance thresholds for using the calculation chain in Aspose.Cells. | Show how to temporarily disable EnableCalculationChain after measurement and recompute formulas to verify identical results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace FastFormulaCalculationDemo
{
    // This C# example builds a workbook with three worksheets, each containing 2,000 rows of numeric data, simple multiplication formulas, and cumulative SUM formulas. It measures the total calculation time using the default engine, then enables the fast formula calculation feature (EnableCalculationChain), recalculates, and records the new timing before saving the file. The sample demonstrates how to profile performance gains when processing large workbooks with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample worksheets and formulas to simulate a heavy workbook
                const int sheetCount = 3;
                const int rowCount = 2000; // number of rows per sheet

                for (int s = 0; s < sheetCount; s++)
                {
                    Worksheet sheet;
                    if (s == 0)
                    {
                        // Use the default first sheet
                        sheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        // Add a new sheet and retrieve it
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    // Populate column A with numeric values
                    for (int r = 0; r < rowCount; r++)
                    {
                        sheet.Cells[r, 0].PutValue(r + 1);
                    }

                    // Populate column B with a simple formula that depends on column A (e.g., =A1*2)
                    for (int r = 0; r < rowCount; r++)
                    {
                        sheet.Cells[r, 1].Formula = $"=A{r + 1}*2";
                    }

                    // Populate column C with a cumulative SUM formula (e.g., =SUM(A1:A{row}))
                    for (int r = 0; r < rowCount; r++)
                    {
                        sheet.Cells[r, 2].Formula = $"=SUM(A1:A{r + 1})";
                    }
                }

                // ------------------------------------------------------------
                // 1. Calculate formulas without calculation chain (default)
                // ------------------------------------------------------------
                Stopwatch sw = new Stopwatch();
                sw.Start();

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                sw.Stop();
                TimeSpan timeWithoutChain = sw.Elapsed;
                Console.WriteLine($"Calculation time without chain: {timeWithoutChain.TotalMilliseconds} ms");

                // ------------------------------------------------------------
                // 2. Enable fast formula calculation (calculation chain)
                // ------------------------------------------------------------
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;

                // Recalculate formulas to measure the effect of the chain
                sw.Restart();

                // First run after enabling may include chain building overhead
                workbook.CalculateFormula();

                sw.Stop();
                TimeSpan timeWithChain = sw.Elapsed;
                Console.WriteLine($"Calculation time with chain enabled: {timeWithChain.TotalMilliseconds} ms");

                // ------------------------------------------------------------
                // Optional: Save the workbook (demonstrates usage of save rule)
                // ------------------------------------------------------------
                workbook.Save("FastFormulaCalculationResult.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
