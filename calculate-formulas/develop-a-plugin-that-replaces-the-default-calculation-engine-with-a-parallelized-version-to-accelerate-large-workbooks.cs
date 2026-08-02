using System;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ParallelCalculationEngineDemo
{
    // Custom calculation engine that processes built‑in functions in parallel
    public class ParallelCalculationEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions (e.g., SUM) by this engine
        public override bool ProcessBuiltInFunctions => true;

        // No special force‑recalculate logic needed
        public override bool ForceRecalculate(string functionName) => false;

        public override void Calculate(CalculationData data)
        {
            // Handle only the SUM function in parallel; other functions fall back to the default engine
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
            {
                double total = 0;
                object sync = new object();

                // Process each parameter concurrently
                Parallel.For(0, data.ParamCount, i =>
                {
                    double localSum = 0;
                    object param = data.GetParamValue(i);

                    // Parameter can be a range (ReferredArea) or a single value
                    if (param is ReferredArea area)
                    {
                        // Iterate over the cells in the range
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                // GetValue expects coordinates relative to the area start
                                object cellVal = area.GetValue(r - area.StartRow, c - area.StartColumn);
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                                    localSum += d;
                            }
                        }
                    }
                    else if (param != null && double.TryParse(param.ToString(), out double d))
                    {
                        localSum = d;
                    }

                    // Accumulate the partial sum safely
                    lock (sync)
                    {
                        total += localSum;
                    }
                });

                // Set the calculated result for the SUM function
                data.CalculatedValue = total;
            }
            // For other functions, do nothing – the default engine will handle them
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load a large workbook (replace with actual path)
            Workbook workbook = new Workbook("LargeWorkbook.xlsx");

            // Instantiate the parallel calculation engine
            var parallelEngine = new ParallelCalculationEngine();

            // Configure calculation options to use the custom engine
            var calcOptions = new CalculationOptions
            {
                CustomEngine = parallelEngine,
                Recursive = true,      // ensure dependent cells are calculated
                IgnoreError = false    // surface any calculation errors
            };

            // Perform formula calculation using the parallel engine
            workbook.CalculateFormula(calcOptions);

            // Save the workbook after calculation
            workbook.Save("LargeWorkbook_ParallelCalculated.xlsx");
        }
    }
}