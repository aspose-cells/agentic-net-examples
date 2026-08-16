// Title: Cache Identical Sub‑Expression Results with a Custom Aspose.Cells Calculation Engine (C#)
// Description: Demonstrates how to extend Aspose.Cells by creating a C# CachingEngine that inherits from AbstractCalculationEngine, builds a unique key from the function name and its parameters, stores computed values in a dictionary, and reuses them for repeated calls. The engine is attached via CalculationOptions.CustomEngine, allowing formulas like =MYFUNC(A1,A2) to be evaluated once even when used in multiple cells, improving performance and reducing redundant calculations.
// Keywords: Aspose.Cells custom engine | C# calculation engine cache | memoization Excel functions | avoid duplicate formula evaluation | performance optimization Aspose.Cells | user‑defined function caching | AbstractCalculationEngine example
// Common Searches: Aspose.Cells custom calculation engine caching | how to memoize custom Excel functions in .NET | reduce repeated formula calculations Aspose.Cells | cache results of user‑defined functions Aspose.Cells | C# Aspose.Cells performance tips for custom functions
// Developer Intent: Create a reusable calculation engine that caches results of identical custom‑function calls to eliminate unnecessary recomputation.
// Use Cases: Speed up large workbooks where the same custom function is invoked many times with identical arguments. | Integrate the CachingEngine into existing Aspose.Cells projects via CalculationOptions to improve overall calculation time. | Extend the caching logic to support multi‑cell ranges or more complex parameter types for advanced user‑defined functions.
// AI Prompts: Generate a thread‑safe version of the CachingEngine that clears its cache when workbook data changes. | Show how to adapt the caching mechanism for custom functions that accept multi‑cell ReferredArea parameters. | Explain step‑by‑step how to attach a custom calculation engine to Aspose.Cells CalculationOptions for optimal performance in spreadsheets with repeated formulas.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCachingEngineDemo
{
    // Custom calculation engine that caches results of identical sub‑expressions
    // Demonstrates how to extend Aspose.Cells by creating a C# CachingEngine that inherits from AbstractCalculationEngine, builds a unique key from the function name and its parameters, stores computed values in a dictionary, and reuses them for repeated calls. The engine is attached via CalculationOptions.CustomEngine, allowing formulas like =MYFUNC(A1,A2) to be evaluated once even when used in multiple cells, improving performance and reducing redundant calculations.
    public class CachingEngine : AbstractCalculationEngine
    {
        // Cache key -> calculated value
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // Build a unique key based on function name and its evaluated parameters
        private string BuildCacheKey(CalculationData data)
        {
            var parts = new List<string> { data.FunctionName.ToUpperInvariant() };
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is ReferredArea area)
                {
                    // For simplicity, use the first cell value of the area
                    object val = area.GetValue(0, 0);
                    parts.Add(val?.ToString() ?? "null");
                }
                else
                {
                    parts.Add(param?.ToString() ?? "null");
                }
            }
            return string.Join("|", parts);
        }

        // Do not force recalculation for our custom function – allow caching
        public override bool ForceRecalculate(string functionName) => false;

        public override void Calculate(CalculationData data)
        {
            // Example custom function name
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                string key = BuildCacheKey(data);

                // Return cached value if present
                if (_cache.TryGetValue(key, out object cachedResult))
                {
                    data.CalculatedValue = cachedResult;
                    return;
                }

                // Compute the result (sum of two parameters)
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);
                    double val = 0;

                    if (param is ReferredArea area)
                    {
                        // Assume single‑cell area for this demo
                        object cellVal = area.GetValue(0, 0);
                        if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                            val = d;
                    }
                    else if (param != null && double.TryParse(param.ToString(), out double d))
                    {
                        val = d;
                    }

                    sum += val;
                }

                // Store result in cache and set it as the calculated value
                _cache[key] = sum;
                data.CalculatedValue = sum;
            }
            // For any other function, let the default engine handle it (do nothing)
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells used by the custom function
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);

            // Two cells using the same sub‑expression (identical parameters)
            sheet.Cells["B1"].Formula = "=MYFUNC(A1,A2)";
            sheet.Cells["B2"].Formula = "=MYFUNC(A1,A2)";

            // Set calculation options with the caching engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachingEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // Perform calculation
            workbook.CalculateFormula(options);

            // Output results – both cells should have the same value, calculated only once
            Console.WriteLine("B1 result: " + sheet.Cells["B1"].Value);
            Console.WriteLine("B2 result: " + sheet.Cells["B2"].Value);

            // Save the workbook (uses the provided save lifecycle)
            workbook.Save("CachingEngineDemo.xlsx");
        }
    }
}
