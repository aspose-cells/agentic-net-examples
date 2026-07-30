// Title: Cache Results in a Custom Aspose.Cells Calculation Engine (C#)
// Description: Shows how to inherit Aspose.Cells' AbstractCalculationEngine and add a dictionary‑based cache that stores the outcome of identical sub‑expressions for a user‑defined function (MYFUNC). The engine creates a case‑insensitive key from the function name and its arguments, returns the cached value on subsequent calls, and is wired through CalculationOptions to accelerate workbook formula evaluation.
// Keywords: Aspose.Cells | C# custom calculation engine | formula result caching | user‑defined function performance | AbstractCalculationEngine example | cached sub‑expression | Excel engine .NET | avoid redundant calculations | MYFUNC sample | performance optimization
// Common Searches: Aspose.Cells custom engine cache example | how to cache formula results in Aspose.Cells | C# user defined function with caching | prevent duplicate calculations in Aspose.Cells | speed up workbook calculations using cache
// Developer Intent: Build a calculation engine that reuses previously computed values for identical sub‑expressions, eliminating unnecessary work.
// Use Cases: Accelerate repeated calls to a custom function across many cells. | Lower CPU usage when large range lookups are evaluated multiple times. | Integrate a cached engine into existing Aspose.Cells workflows for high‑performance reporting.
// AI Prompts: Create a C# version of CachedCalculationEngine that supports multiple user‑defined functions with independent caches. | Explain how to extend the cache key to include full range addresses for multi‑cell parameters. | Write unit tests that verify cache hits, misses, and thread‑safety for the custom engine.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomEngineDemo
{
    // Custom calculation engine that caches results of identical sub‑expressions
    // Shows how to inherit Aspose.Cells' AbstractCalculationEngine and add a dictionary‑based cache that stores the outcome of identical sub‑expressions for a user‑defined function (MYFUNC). The engine creates a case‑insensitive key from the function name and its arguments, returns the cached value on subsequent calls, and is wired through CalculationOptions to accelerate workbook formula evaluation.
    public class CachedCalculationEngine : AbstractCalculationEngine
    {
        // Simple cache: key -> calculated value
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // No need to force recalculation for cached functions
        public override bool ForceRecalculate(string functionName) => false;

        // Core calculation method
        public override void Calculate(CalculationData data)
        {
            // Example custom function name
            const string targetFunction = "MYFUNC";

            // Only handle our custom function; let default engine process others
            if (!data.FunctionName.Equals(targetFunction, StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key based on function name and parameter values
            var keyParts = new List<string> { targetFunction };
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                // If the parameter is a range, extract its first cell value
                if (param is ReferredArea area)
                {
                    param = area.GetValue(0, 0);
                }
                keyParts.Add(Convert.ToString(param));
            }
            string cacheKey = string.Join("|", keyParts);

            // Return cached result if present
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // Perform the actual calculation (sum of two parameters in this example)
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is ReferredArea area)
                {
                    param = area.GetValue(0, 0);
                }
                sum += Convert.ToDouble(param);
            }

            // Store and return the result
            data.CalculatedValue = sum;
            _cache[cacheKey] = sum;
        }
    }

    class Program
    {
        static void Main()
        {
            // ==== Create a new workbook ====
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);

            // Formulas using the custom function MYFUNC
            // Both formulas are identical; the second should hit the cache
            sheet.Cells["C1"].Formula = "=MYFUNC(A1, B1)";
            sheet.Cells["C2"].Formula = "=MYFUNC(A1, B1)";

            // A different sub‑expression (different parameters) – not cached yet
            sheet.Cells["C3"].Formula = "=MYFUNC(A2, B2)";

            // ==== Set calculation options with the custom engine ====
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachedCalculationEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // ==== Calculate all formulas ====
            workbook.CalculateFormula(options);

            // ==== Output results ====
            Console.WriteLine("C1 (cached result): " + sheet.Cells["C1"].Value); // Expected 15
            Console.WriteLine("C2 (should use cache): " + sheet.Cells["C2"].Value); // Expected 15
            Console.WriteLine("C3 (different parameters): " + sheet.Cells["C3"].Value); // Expected 45

            // ==== Save the workbook ====
            workbook.Save("CachedEngineDemo.xlsx");
        }
    }
}
