// Title: Implement a parallel custom calculation engine for Aspose.Cells to speed up SUM formulas in large Excel workbooks (C#)
// AI Prompts: Write a C# class that derives from Aspose.Cells.AbstractCalculationEngine and uses Parallel.For to compute the SUM function across cell ranges. | Show how to assign the ParallelCalculationEngine to CalculationOptions and call Workbook.CalculateFormula on a large workbook. | Provide code that extends the ParallelCalculationEngine to process AVERAGE and MAX functions concurrently using parallel execution.
// Common Searches: how to implement a multithreaded SUM function in Aspose.Cells using Parallel.For | replace default formula engine with a custom parallel engine in Aspose.Cells .NET | optimize Workbook.CalculateFormula performance for large Excel files in C#
// Tags: parallel formula engine Aspose.Cells C# | multithreaded SUM calculation with Parallel.For | configure calculation options for custom engine | large workbook formula performance optimization | extend engine to support AVERAGE and MAX

using System;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ParallelCalculationEnginePlugin
{
    // Custom calculation engine that parallelizes the processing of built‑in functions (e.g., SUM)
    // The example defines a ParallelCalculationEngine that overrides Aspose.Cells.AbstractCalculationEngine to calculate the SUM function in parallel with Parallel.For, configures CalculationOptions.CustomEngine to use this engine, runs Workbook.CalculateFormula on a large workbook, and saves the calculated result.
    public class ParallelCalculationEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so that this engine gets called for them
        public override bool ProcessBuiltInFunctions => true;

        // No special force‑recalculation logic needed
        public override bool ForceRecalculate(string functionName) => false;

        // Core calculation method
        public override void Calculate(CalculationData data)
        {
            // Example: parallel implementation for the SUM function
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double total = 0.0;
                object sync = new object();

                // Iterate over each parameter (each can be a range or a scalar)
                for (int p = 0; p < data.ParamCount; p++)
                {
                    object param = data.GetParamValue(p);

                    // If the parameter is a range (ReferredArea), sum its cells in parallel
                    if (param is ReferredArea area)
                    {
                        Parallel.For(area.StartRow, area.EndRow + 1, row =>
                        {
                            double rowSum = 0.0;
                            for (int col = area.StartColumn; col <= area.EndColumn; col++)
                            {
                                // GetValue expects zero‑based indices relative to the area
                                object cellVal = area.GetValue(row - area.StartRow, col - area.StartColumn);
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                                    rowSum += d;
                            }
                            // Accumulate row sums safely
                            lock (sync) { total += rowSum; }
                        });
                    }
                    // If the parameter is a scalar value, add it directly
                    else if (param != null && double.TryParse(param.ToString(), out double scalar))
                    {
                        total += scalar;
                    }
                }

                // Set the calculated result for the SUM function
                data.CalculatedValue = total;
                return;
            }

            // For other functions, let the default engine handle them (do nothing here)
        }
    }

    public class PluginDemo
    {
        public static void Run()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("LargeWorkbook.xlsx");

            // Configure calculation options to use the parallel engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new ParallelCalculationEngine(),
                // Keep other options as needed
                IgnoreError = false,
                Recursive = true
            };

            // Perform formula calculation using the custom parallel engine
            workbook.CalculateFormula(options);

            // Save the workbook after calculation (replace with desired output path)
            workbook.Save("LargeWorkbook_Calculated.xlsx");
        }
    }

    // Entry point for testing
    class Program
    {
        static void Main()
        {
            PluginDemo.Run();
        }
    }
}
