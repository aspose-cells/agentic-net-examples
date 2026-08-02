// Title: Speed up Aspose.Cells Cell.Calculate with a C# caching custom calculation engine for MYHEAVYFUNC
// Description: Demonstrates how to derive a CachingEngine from AbstractCalculationEngine, cache the result of a non‑volatile custom function (MYHEAVYFUNC), and plug it into CalculationOptions to cut the runtime of wb.CalculateFormula dramatically. The example fills 1,000 identical formulas, measures execution time, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | custom calculation engine | AbstractCalculationEngine | caching | performance optimization | Cell.Calculate | formula engine | heavy custom function | reduce execution time | GitHub example
// Common Searches: Aspose.Cells cache custom function result | speed up Cell.Calculate in C# | AbstractCalculationEngine example | optimize heavy formula Aspose.Cells | prevent recalculation of non‑volatile functions | Aspose.Cells performance tips
// Developer Intent: Create a custom calculation engine that stores and reuses the output of a costly, non‑volatile function to avoid repeated work during workbook recalculation.
// Use Cases: Reuse the same MYHEAVYFUNC result across thousands of cells, executing the expensive loop only once. | Turn off built‑in function processing when the engine handles only custom functions, lowering overhead. | Keep cached values when the workbook is recalculated by returning false from ForceRecalculate.
// AI Prompts: Generate a thread‑safe CachingEngine using ConcurrentDictionary with optional expiration. | Extend the engine to support multiple heavy custom functions, each with its own cache. | Write NUnit tests that verify cache hits, cache misses, and compare performance against the default engine.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsOptimizationDemo
{
    // Custom engine that caches results of a heavy custom function.
    // Demonstrates how to derive a CachingEngine from AbstractCalculationEngine, cache the result of a non‑volatile custom function (MYHEAVYFUNC), and plug it into CalculationOptions to cut the runtime of wb.CalculateFormula dramatically. The example fills 1,000 identical formulas, measures execution time, and saves the workbook.
    public class CachingEngine : AbstractCalculationEngine
    {
        // Cache key is a string representation of all parameters.
        private static readonly Dictionary<string, object> _cache = new Dictionary<string, object>();

        // The custom function is not volatile, so we can reuse cached results.
        public override bool ForceRecalculate(string functionName)
        {
            return false; // Do not force recalculation for MYHEAVYFUNC
        }

        // No need to process built‑in functions; keep it false for performance.
        public override bool ProcessBuiltInFunctions => false;

        public override void Calculate(CalculationData data)
        {
            // Handle only our custom function.
            if (!string.Equals(data.FunctionName, "MYHEAVYFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let the default engine handle other functions.

            // Build a cache key from the literal text of each parameter.
            var keyBuilder = new StringBuilder();
            for (int i = 0; i < data.ParamCount; i++)
            {
                string paramText = data.GetParamText(i) ?? string.Empty;
                keyBuilder.Append(paramText);
                keyBuilder.Append('|');
            }
            string cacheKey = keyBuilder.ToString();

            // Try to get a cached result.
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // ----- Expensive calculation starts -----
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                // Get the parameter value as a double.
                object param = data.GetParamValue(i);
                double val;

                // If the parameter is a range, Aspose.Cells returns a 2‑D object array.
                if (param is object[,] arr && arr.Length > 0)
                {
                    // Use the first cell of the range.
                    val = Convert.ToDouble(arr.GetValue(0, 0) ?? 0);
                }
                else
                {
                    val = Convert.ToDouble(param ?? 0);
                }

                // Simulate work per parameter.
                for (int j = 0; j < 10000; j++)
                {
                    sum += Math.Sqrt(val + j);
                }
            }
            // ----- Expensive calculation ends -----

            // Store the result in the cache for future reuse.
            _cache[cacheKey] = sum;
            data.CalculatedValue = sum;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook and fill data.
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Populate input values.
                for (int i = 0; i < 10; i++)
                    cells[i, 0].PutValue(i + 1); // Column A

                // Apply the custom function to many cells.
                // All cells share the same parameters, so caching will help.
                for (int row = 0; row < 1000; row++)
                    cells[row, 1].Formula = "=MYHEAVYFUNC(A1,A2)";

                // Set calculation options with the custom engine.
                CalculationOptions opts = new CalculationOptions
                {
                    CustomEngine = new CachingEngine(),
                    Recursive = true,
                    IgnoreError = false
                };

                // Measure calculation time.
                Stopwatch sw = Stopwatch.StartNew();
                wb.CalculateFormula(opts);
                sw.Stop();

                Console.WriteLine($"Total calculation time: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine($"Sample result (B1): {cells[0, 1].Value}");

                // Save the workbook.
                string outputPath = "OptimizedResult.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
