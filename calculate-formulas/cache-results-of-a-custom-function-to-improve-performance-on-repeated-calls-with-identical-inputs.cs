// Title: Cache Custom Function Results in Aspose.Cells for .NET – Example Using AbstractCalculationEngine
// Description: Demonstrates how to create a CachedEngine that inherits AbstractCalculationEngine, overrides ForceRecalculate, and stores results of the custom function MYCACHEDFUNC in a Dictionary. The sample shows building a unique cache key from two numeric parameters, returning cached values on repeated calls, and falling back to calculation when inputs change. It also illustrates setting CalculationOptions.CustomEngine, running workbook.CalculateFormula, modifying source cells, and saving the workbook, providing a clear performance boost for non‑volatile custom formulas.
// Keywords: Aspose.Cells | custom function cache | C# | .NET | AbstractCalculationEngine | formula performance | Excel calculation engine | cache dictionary | non‑volatile function | GitHub example | code snippet
// Common Searches: Aspose.Cells cache custom function results | How to implement caching in Aspose.Cells calculation engine | Avoid recalculating custom formulas in .NET Excel library | Performance optimization for Aspose.Cells custom functions | Example of AbstractCalculationEngine with caching
// Developer Intent: Implement a caching layer for a custom Aspose.Cells function to eliminate redundant calculations and improve workbook processing speed.
// Use Cases: Speed up large spreadsheets that call the same custom function with identical arguments. | Maintain cached values after minor data edits so only cells with changed inputs are recomputed. | Apply caching to complex, non‑volatile custom formulas such as lookups, statistical models, or financial calculations.
// AI Prompts: Generate a thread‑safe version of CachedEngine that works with parallel calculations in Aspose.Cells. | Extend the cache to support any number of parameters and mixed data types for custom functions. | Create unit tests for CachedEngine that verify cache hits, cache misses, and correct handling of parameter changes.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionCacheDemo
{
    // Custom calculation engine that caches results of MYCACHEDFUNC
    // Demonstrates how to create a CachedEngine that inherits AbstractCalculationEngine, overrides ForceRecalculate, and stores results of the custom function MYCACHEDFUNC in a Dictionary. The sample shows building a unique cache key from two numeric parameters, returning cached values on repeated calls, and falling back to calculation when inputs change. It also illustrates setting CalculationOptions.CustomEngine, running workbook.CalculateFormula, modifying source cells, and saving the workbook, providing a clear performance boost for non‑volatile custom formulas.
    public class CachedEngine : AbstractCalculationEngine
    {
        // Simple cache: key is a string representation of the parameters, value is the calculated result
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>();

        // Do not force recalculation for the custom function – allow cached values to be reused
        public override bool ForceRecalculate(string functionName)
        {
            return false; // return true only for volatile functions
        }

        // Core calculation logic
        public override void Calculate(CalculationData data)
        {
            // Handle only our custom function
            if (!data.FunctionName.Equals("MYCACHEDFUNC", StringComparison.OrdinalIgnoreCase))
                return; // let the default engine handle other functions

            // Retrieve parameter values (assume two numeric parameters)
            object param0 = data.GetParamValue(0);
            object param1 = data.GetParamValue(1);

            // Build a cache key that uniquely identifies the input combination
            string key = $"{param0}_{param1}";

            // If we have a cached result, reuse it
            if (_cache.TryGetValue(key, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // Otherwise perform the actual calculation (example: multiply the two numbers)
            double val0 = Convert.ToDouble(param0);
            double val1 = Convert.ToDouble(param1);
            double result = val0 * val1; // custom logic

            // Store the result in the cache for future reuse
            _cache[key] = result;

            // Return the calculated value to the engine
            data.CalculatedValue = result;
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(2);
            cells["B1"].PutValue(3);
            cells["A2"].PutValue(2);
            cells["B2"].PutValue(3);
            cells["A3"].PutValue(5);
            cells["B3"].PutValue(7);

            // Use the custom function in several cells with identical and different inputs
            cells["C1"].Formula = "=MYCACHEDFUNC(A1,B1)"; // first unique set (2,3)
            cells["C2"].Formula = "=MYCACHEDFUNC(A2,B2)"; // same as C1 – should hit cache
            cells["C3"].Formula = "=MYCACHEDFUNC(A3,B3)"; // different set (5,7)

            // ---------- Set calculation options with the custom engine ----------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachedEngine()
            };

            // First calculation – cache will be populated
            workbook.CalculateFormula(options);
            Console.WriteLine($"C1 = {cells["C1"].Value} (expected 6)");
            Console.WriteLine($"C2 = {cells["C2"].Value} (expected 6, from cache)");
            Console.WriteLine($"C3 = {cells["C3"].Value} (expected 35)");

            // Change a source cell to see that cached result is reused for unchanged inputs
            cells["A1"].PutValue(2); // same value, cache still valid
            cells["A3"].PutValue(10); // change input for C3, will recalculate

            // Second calculation – C1 and C2 should use cached value, C3 recalculates
            workbook.CalculateFormula(options);
            Console.WriteLine($"After modification:");
            Console.WriteLine($"C1 = {cells["C1"].Value} (cached 6)");
            Console.WriteLine($"C2 = {cells["C2"].Value} (cached 6)");
            Console.WriteLine($"C3 = {cells["C3"].Value} (recalculated 70)");

            // ---------- Save the workbook ----------
            workbook.Save("CustomFunctionCacheDemo.xlsx");
        }
    }
}
