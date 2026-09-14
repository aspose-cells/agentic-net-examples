// Title: Cache results of a custom Excel function using Aspose.Cells C# AbstractCalculationEngine
// AI Prompts: Create an AbstractCalculationEngine subclass that records MYCACHEFUNC outputs in a Dictionary and returns the cached value during workbook.CalculateFormula. | Generate a unique cache key from scalar and single‑cell range arguments, handling nulls and complex types gracefully. | Wire the custom engine into CalculationOptions, run identical formulas, modify an unrelated cell, and confirm that cached results are reused without recomputation.
// Common Searches: Aspose.Cells how to implement caching for user defined functions in C# | example of AbstractCalculationEngine with in‑memory dictionary cache | prevent recalculation of custom Excel functions using Aspose.Cells calculation options | performance tip for repeated MYCACHEFUNC calls in Aspose.Cells workbook
// Tags: custom calculation engine cache | user defined function memoization | Aspose.Cells in‑memory cache | calculation options custom engine | optimize repeated formula evaluation

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionCacheDemo
{
    // Custom calculation engine that caches results of the custom function "MYCACHEFUNC"
    // The sample defines a CachingEngine class that inherits from AbstractCalculationEngine, builds a string key from MYCACHEFUNC parameters, checks a Dictionary cache for a pre‑computed sum, stores new results, and assigns the cached value to the calculation data. The program applies the function to worksheet cells with identical inputs, changes an unrelated cell, recalculates to demonstrate that cached values persist, and saves the workbook.
    public class CachingEngine : AbstractCalculationEngine
    {
        // Simple in‑memory cache: key = concatenated parameter values, value = calculated result
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // The engine does not need parameters in array mode
        public override bool IsParamArrayModeRequired => false;

        // Do not force recalculation for the custom function – allow caching
        public override bool ForceRecalculate(string functionName)
        {
            return false; // return true only for volatile functions
        }

        // Main calculation method
        public override void Calculate(CalculationData data)
        {
            // We only handle our custom function; other functions fall back to the default engine
            if (!string.Equals(data.FunctionName, "MYCACHEFUNC", StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key from all parameter values
            var keyParts = new List<string>();
            for (int i = 0; i < data.ParamCount; i++)
            {
                // Get the raw parameter value (could be a scalar, ReferredArea, etc.)
                object param = data.GetParamValue(i);

                // For simplicity, handle scalar values and single‑cell ReferredArea
                if (param is double d)
                {
                    keyParts.Add(d.ToString());
                }
                else if (param is ReferredArea area && area.StartRow == area.EndRow && area.StartColumn == area.EndColumn)
                {
                    // Single cell – fetch its value
                    object cellVal = area.GetValue(0, 0);
                    keyParts.Add(cellVal?.ToString() ?? "null");
                }
                else
                {
                    // Fallback for complex types – use their string representation
                    keyParts.Add(param?.ToString() ?? "null");
                }
            }

            string cacheKey = string.Join("|", keyParts);

            // Check cache
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                // Use cached value
                data.CalculatedValue = cachedResult;
                return;
            }

            // Perform the actual calculation (example: sum of all numeric parameters)
            double sum = 0;
            foreach (string part in keyParts)
            {
                if (double.TryParse(part, out double val))
                    sum += val;
            }

            // Store result in cache and set it as the calculated value
            _cache[cacheKey] = sum;
            data.CalculatedValue = sum;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(5);
            cells["B2"].PutValue(15);

            // Use the custom function in several cells with identical inputs
            cells["C1"].Formula = "=MYCACHEFUNC(A1, B1)"; // 10 + 5 = 15
            cells["C2"].Formula = "=MYCACHEFUNC(A1, B1)"; // same inputs – should hit cache
            cells["C3"].Formula = "=MYCACHEFUNC(A2, B2)"; // 20 + 15 = 35
            cells["C4"].Formula = "=MYCACHEFUNC(A2, B2)"; // same inputs – should hit cache

            // Set up calculation options with the custom caching engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachingEngine()
            };

            // First calculation – cache will be populated
            workbook.CalculateFormula(options);
            Console.WriteLine($"C1 = {cells["C1"].Value} (expected 15)");
            Console.WriteLine($"C2 = {cells["C2"].Value} (expected 15, cached)");
            Console.WriteLine($"C3 = {cells["C3"].Value} (expected 35)");
            Console.WriteLine($"C4 = {cells["C4"].Value} (expected 35, cached)");

            // Change an unrelated cell (does not affect cached function)
            cells["D1"].PutValue(999);

            // Second calculation – cached results should be reused without recomputation
            workbook.CalculateFormula(options);
            Console.WriteLine("After modifying unrelated cell D1:");
            Console.WriteLine($"C1 = {cells["C1"].Value}");
            Console.WriteLine($"C2 = {cells["C2"].Value}");
            Console.WriteLine($"C3 = {cells["C3"].Value}");
            Console.WriteLine($"C4 = {cells["C4"].Value}");

            // Save the workbook
            workbook.Save("CachingEngineDemo.xlsx");
        }
    }
}
