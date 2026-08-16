// Title: Parallel Custom Calculation Engine for Aspose.Cells (C#) – Accelerate Large Workbook Formulas
// Description: Demonstrates a ParallelCalculationEngine derived from AbstractCalculationEngine that processes the SUM function in parallel using Parallel.For, integrates via CalculationOptions.CustomEngine, and speeds up formula evaluation on a 10,000‑row worksheet.
// Keywords: Aspose.Cells | custom calculation engine | parallel formula evaluation | C# | .NET | performance optimization | large workbook | SUM function parallelism | CalculationOptions.CustomEngine | multithreaded Excel processing
// Common Searches: Aspose.Cells custom calculation engine example | parallel formula calculation C# Aspose.Cells | replace default calculation engine Aspose | speed up SUM formula in large workbook | multithreaded Excel calculation .NET
// Developer Intent: Implement a custom parallel calculation engine to replace Aspose.Cells' default engine and improve formula performance in large workbooks.
// Use Cases: Compute SUM over tens of thousands of rows with Parallel.For to reduce calculation time. | Plug the ParallelCalculationEngine into CalculationOptions.CustomEngine for whole‑workbook formula evaluation. | Fallback to the built‑in engine for unsupported functions while accelerating supported ones.
// AI Prompts: Create a ParallelCalculationEngine that also handles AVERAGE and MAX with thread‑safe aggregation. | Show a benchmark comparing the default engine and the parallel engine on a 50,000‑row workbook. | Add robust error handling for non‑numeric cells in ReferredArea when using the parallel engine.

using System;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ParallelCalculationEngineDemo
{
    // Custom calculation engine that processes built‑in functions in parallel
    // Demonstrates a ParallelCalculationEngine derived from AbstractCalculationEngine that processes the SUM function in parallel using Parallel.For, integrates via CalculationOptions.CustomEngine, and speeds up formula evaluation on a 10,000‑row worksheet.
    public class ParallelCalculationEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions (e.g., SUM)
        public override bool ProcessBuiltInFunctions => true;

        // No forced recalculation for any function
        public override bool ForceRecalculate(string functionName) => false;

        // Core calculation logic
        public override void Calculate(CalculationData data)
        {
            // Example: parallel implementation for the SUM function
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double totalSum = 0.0;
                object syncRoot = new object();

                // Iterate over each parameter passed to SUM
                for (int p = 0; p < data.ParamCount; p++)
                {
                    // Parameters are usually ReferredArea objects (ranges)
                    if (data.GetParamValue(p) is ReferredArea area)
                    {
                        // Parallel loop over rows in the area
                        Parallel.For(area.StartRow, area.EndRow + 1, row =>
                        {
                            double rowSum = 0.0;
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                object cellVal = area.GetValue(row - area.StartRow, col - area.StartColumn);
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                                {
                                    rowSum += d;
                                }
                            }
                            // Accumulate safely
                            lock (syncRoot)
                            {
                                totalSum += rowSum;
                            }
                        });
                    }
                    else
                    {
                        // Handle scalar parameters (e.g., numbers)
                        object val = data.GetParamValue(p);
                        if (val != null && double.TryParse(val.ToString(), out double d))
                        {
                            totalSum += d;
                        }
                    }
                }

                // Set the calculated result for the SUM function
                data.CalculatedValue = totalSum;
            }
            // For other functions, let the default engine handle them
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large range of data to demonstrate parallel calculation
            const int rowCount = 10000;
            for (int i = 0; i < rowCount; i++)
            {
                // Fill column A with incremental numbers
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // Set a formula that sums the entire column A
            sheet.Cells["B1"].Formula = $"=SUM(A1:A{rowCount})";

            // Configure calculation options with the custom parallel engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new ParallelCalculationEngine(),
                Recursive = true,
                IgnoreError = false
            };

            // Calculate all formulas using the custom engine (lifecycle: calculate)
            workbook.CalculateFormula(calcOptions);

            // Output the result to console for verification
            Console.WriteLine("Result of SUM(A1:A{0}) = {1}", rowCount, sheet.Cells["B1"].Value);

            // Save the workbook (lifecycle: save)
            workbook.Save("ParallelCalculationResult.xlsx");
        }
    }
}
