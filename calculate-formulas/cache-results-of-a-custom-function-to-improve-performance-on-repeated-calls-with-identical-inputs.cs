using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionCaching
{
    // Custom calculation engine that caches results of MYCACHEDFUNC
    public class CachingEngine : AbstractCalculationEngine
    {
        // Simple in‑memory cache: key = concatenated parameter values, value = calculated result
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // Do not force recalculation – allow the engine to reuse cached values
        public override bool ForceRecalculate(string functionName) => false;

        public override void Calculate(CalculationData data)
        {
            // Only handle our custom function
            if (!data.FunctionName.Equals("MYCACHEDFUNC", StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key from all parameters
            StringBuilder keyBuilder = new StringBuilder();
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);

                // If the parameter is a range (ReferredArea), use its first cell value
                if (param is ReferredArea area)
                {
                    object val = area.GetValue(0, 0);
                    keyBuilder.Append(val?.ToString() ?? "null");
                }
                else
                {
                    keyBuilder.Append(param?.ToString() ?? "null");
                }

                keyBuilder.Append("|"); // separator
            }

            string cacheKey = keyBuilder.ToString();

            // Return cached result if it exists
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // Example calculation: sum of all numeric parameters
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);

                if (param is ReferredArea area)
                {
                    object val = area.GetValue(0, 0);
                    if (val != null)
                        sum += Convert.ToDouble(val);
                }
                else if (param != null)
                {
                    sum += Convert.ToDouble(param);
                }
            }

            // Store result in cache and set it as the function result
            data.CalculatedValue = sum;
            _cache[cacheKey] = sum;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some input cells
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["A2"].PutValue(10); // same as A1 to demonstrate caching
            cells["B2"].PutValue(20); // same as B1

            // Use the custom function in two different cells with identical inputs
            cells["C1"].Formula = "=MYCACHEDFUNC(A1,B1)";
            cells["C2"].Formula = "=MYCACHEDFUNC(A2,B2)";

            // Set calculation options to use our custom engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachingEngine()
            };

            // First calculation – engine computes the result and caches it
            workbook.CalculateFormula(options);
            Console.WriteLine($"First calculation C1: {cells["C1"].Value}"); // Expected 30
            Console.WriteLine($"First calculation C2: {cells["C2"].Value}"); // Expected 30 (cached)

            // Change a cell that is NOT used by the function to prove caching still works
            cells["D1"].PutValue(999);

            // Second calculation – cached value should be reused, no recomputation
            workbook.CalculateFormula(options);
            Console.WriteLine($"Second calculation C1: {cells["C1"].Value}");
            Console.WriteLine($"Second calculation C2: {cells["C2"].Value}");

            // Save the workbook (demonstrates required save rule)
            workbook.Save("CachingCustomFunctionDemo.xlsx");
        }
    }
}