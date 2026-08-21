// Title: ClearCache method for resetting a custom calculation engine cache in Aspose.Cells .NET
// Description: Shows how to implement a public ClearCache() method in a CachingCustomEngine that inherits AbstractCalculationEngine. The method clears the internal Dictionary that stores results of custom functions such as MYSUM, so after changing source cells a fresh wb.CalculateFormula call returns updated values.
// Keywords: Aspose.Cells | .NET | C# | custom calculation engine | cache clearing | ClearCache method | MYSUM function | recalculate formulas | dictionary cache | Excel automation | US developers | global
// Common Searches: how to clear cache in Aspose.Cells custom engine | reset custom function results before recalculation .NET | Aspose.Cells C# clear internal cache of calculation engine | force formula recompute after data change Aspose.Cells | C# example for clearing custom engine cache in Aspose.Cells
// Developer Intent: Provide a simple public method that empties the custom engine’s cache so formulas are recomputed after any data modifications.
// Use Cases: Call engine.ClearCache() after user edits a cell, then recalculate the workbook to display correct custom function results. | Integrate ClearCache into a batch‑processing routine that updates many worksheets before final calculation. | Invoke ClearCache before saving a workbook to guarantee that all cached custom function values are up‑to‑date.
// AI Prompts: Generate a C# ClearCache() implementation for a class derived from AbstractCalculationEngine that uses a Dictionary cache. | Write unit tests that verify ClearCache removes all cached entries and forces a new calculation of the MYSUM custom function. | Create a step‑by‑step guide on incorporating ClearCache into an Aspose.Cells workbook recalculation workflow with CalculationOptions.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace CustomEngineDemo
{
    // Custom calculation engine that caches results of custom functions.
    // Shows how to implement a public ClearCache() method in a CachingCustomEngine that inherits AbstractCalculationEngine. The method clears the internal Dictionary that stores results of custom functions such as MYSUM, so after changing source cells a fresh wb.CalculateFormula call returns updated values.
    public class CachingCustomEngine : AbstractCalculationEngine
    {
        // Simple in‑memory cache: key = function name + parameters, value = calculated result.
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>();

        // Public method to clear the cache. Call this after data changes to force fresh computation.
        public void ClearCache()
        {
            _cache.Clear();
        }

        // Example implementation of a custom function "MYSUM".
        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase))
            {
                // Build a cache key based on the function name and its parameters.
                string cacheKey = BuildCacheKey(data);

                // If we have a cached value, reuse it.
                if (_cache.TryGetValue(cacheKey, out object cachedResult))
                {
                    data.CalculatedValue = cachedResult;
                    return;
                }

                // Otherwise compute the result.
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    // Get each parameter as a ReferredArea (range or single cell).
                    ReferredArea area = (ReferredArea)data.GetParamValue(i);
                    // For simplicity, assume each area is a single cell.
                    object val = area.GetValue(0, 0);
                    if (val != null && double.TryParse(val.ToString(), out double d))
                    {
                        sum += d;
                    }
                }

                // Store the result in the cache and set it as the calculated value.
                _cache[cacheKey] = sum;
                data.CalculatedValue = sum;
            }
            else
            {
                // For all other functions let the default engine handle them.
                // No action needed because this method is abstract; simply do nothing.
            }
        }

        // Helper to create a deterministic cache key.
        private string BuildCacheKey(CalculationData data)
        {
            var parts = new List<string> { data.FunctionName.ToUpperInvariant() };
            for (int i = 0; i < data.ParamCount; i++)
            {
                ReferredArea area = (ReferredArea)data.GetParamValue(i);
                // Include the address of the area; for a single cell this is enough.
                parts.Add($"{area.StartRow}:{area.StartColumn}-{area.EndRow}:{area.EndColumn}");
            }
            return string.Join("|", parts);
        }

        // Force recalculation for the custom function so that shared formulas are evaluated per cell.
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYSUM", StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class Program
    {
        public static void Main()
        {
            // Create a workbook and fill some data.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].Formula = "=MYSUM(A1,A2)";

            // Instantiate the custom engine.
            var engine = new CachingCustomEngine();

            // Set calculation options to use the custom engine.
            CalculationOptions opts = new CalculationOptions { CustomEngine = engine };

            // First calculation – result will be computed and cached.
            wb.CalculateFormula(opts);
            Console.WriteLine($"First result: {ws.Cells["A3"].Value}"); // Expected 30

            // Change one of the source cells.
            ws.Cells["A1"].PutValue(100);

            // Without clearing the cache the old result would be returned.
            wb.CalculateFormula(opts);
            Console.WriteLine($"Result without clearing cache: {ws.Cells["A3"].Value}"); // Still 30 (cached)

            // Clear the engine's cache to force fresh computation.
            engine.ClearCache();

            // Re‑calculate after clearing cache – now the new value is used.
            wb.CalculateFormula(opts);
            Console.WriteLine($"Result after clearing cache: {ws.Cells["A3"].Value}"); // Expected 120

            // Save the workbook if needed.
            wb.Save("CustomEngineCacheDemo.xlsx");
        }
    }
}
