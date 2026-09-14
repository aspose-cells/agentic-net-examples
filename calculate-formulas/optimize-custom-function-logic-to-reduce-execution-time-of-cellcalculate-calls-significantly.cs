// Title: Create a cache‑enabled FASTSUM custom function in Aspose.Cells for .NET to accelerate repeated Cell.Calculate evaluations
// AI Prompts: Write a custom Aspose.Cells calculation engine that intercepts the FASTSUM function, computes the sum once per range, and returns the cached value on subsequent calls. | Design a method to generate a unique identifier for a ReferredArea range and store its sum in a Dictionary for fast lookup. | Integrate the custom engine with Workbook.CalculateFormula on a sheet containing thousands of identical FASTSUM formulas and compare execution times.
// Common Searches: asp.net aspocells cache result of custom FASTSUM function | optimize Cell.Calculate performance when many formulas use the same range | example of custom calculation engine with caching in Aspose.Cells C# | reduce execution time of repeated custom functions in Aspose.Cells workbook
// Tags: custom calculation engine caching Aspose.Cells | FASTSUM function performance optimization C# | range sum cache dictionary Aspose.Cells | bulk formula evaluation speed Aspose.Cells | Cell.Calculate acceleration with custom engine

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample builds a workbook, fills column A with 10,000 numbers, adds 1,000 FASTSUM formulas that reference the same range, and defines a FastSumEngine that caches each range's sum using a dictionary keyed by the range coordinates. The custom engine is supplied to CalculationOptions, enabling Workbook.CalculateFormula to reuse cached results and dramatically cut execution time before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a large range with numeric values (A1:A10000)
            for (int i = 0; i < 10000; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // Insert many formulas that call the custom function FASTSUM over the same range.
            // Without optimization each call would recompute the sum, causing a huge slowdown.
            for (int i = 0; i < 1000; i++)
            {
                sheet.Cells[i, 1].Formula = "=FASTSUM(A1:A10000)";
            }

            // Configure calculation options to use the custom engine.
            // CalcStackSize is left at default (200) – adjusting it can help with deep dependency trees.
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new FastSumEngine()
            };

            // Calculate all formulas in the workbook using the custom engine.
            workbook.CalculateFormula(options);

            // Display a sample result to verify correctness.
            Console.WriteLine("Result of FASTSUM: " + sheet.Cells[0, 1].Value);

            // Save the workbook using the standard Aspose.Cells API (required lifecycle rule).
            string outputPath = "OptimizedFastSum.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Custom calculation engine that implements a fast, cache‑aware FASTSUM function.
    class FastSumEngine : AbstractCalculationEngine
    {
        // Cache stores previously computed sums keyed by a unique range identifier.
        private readonly Dictionary<string, double> _cache = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        // FASTSUM does not need to be forced to recalculate for each cell; reuse cached results.
        public override bool ForceRecalculate(string functionName) => false;

        // We are not overriding any built‑in functions, so keep this false for best performance.
        public override bool ProcessBuiltInFunctions => false;

        public override void Calculate(CalculationData data)
        {
            // Handle only the custom function FASTSUM; other functions fall back to the default engine.
            if (!string.Equals(data.FunctionName, "FASTSUM", StringComparison.OrdinalIgnoreCase))
                return;

            // FASTSUM expects a single parameter that is a range (ReferredArea).
            object param = data.GetParamValue(0);
            string cacheKey = GetRangeKey(param);

            // Return cached result if the same range was already summed.
            if (_cache.TryGetValue(cacheKey, out double cachedValue))
            {
                data.CalculatedValue = cachedValue;
                return;
            }

            double sum = 0;

            // Compute the sum manually by iterating over the cells in the range.
            if (param is ReferredArea area)
            {
                for (int r = area.StartRow; r <= area.EndRow; r++)
                {
                    for (int c = area.StartColumn; c <= area.EndColumn; c++)
                    {
                        // GetValue expects coordinates relative to the start of the area.
                        object cellValue = area.GetValue(r - area.StartRow, c - area.StartColumn);
                        if (cellValue is double d)
                            sum += d;
                        else if (cellValue is int i)
                            sum += i;
                        else if (cellValue != null && double.TryParse(cellValue.ToString(), out double parsed))
                            sum += parsed;
                    }
                }
            }

            // Store the computed sum in the cache for future reuse.
            _cache[cacheKey] = sum;
            data.CalculatedValue = sum;
        }

        // Generates a unique string key for a range parameter to be used in the cache.
        private string GetRangeKey(object param)
        {
            if (param is ReferredArea area)
            {
                // Use row/column indices to uniquely identify the range (worksheet name omitted as only one sheet is used).
                return $"R{area.StartRow}_C{area.StartColumn}_R{area.EndRow}_C{area.EndColumn}";
            }
            // Fallback for scalar parameters.
            return param?.ToString() ?? string.Empty;
        }
    }
}
