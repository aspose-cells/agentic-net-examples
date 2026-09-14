// Title: Create a C# Aspose.Cells custom calculation engine that caches identical sub‑expression results
// AI Prompts: Write a C# class inheriting AbstractCalculationEngine that generates a cache key from the function name and its arguments and stores computed values in a Dictionary. | Set up CalculationOptions to use the custom engine, then invoke Workbook.CalculateFormula to demonstrate cache hits for repeated MYCACHEFUNC calls. | Extend the caching logic to handle ReferredArea parameters so that identical cell range values also retrieve results from the cache.
// Common Searches: Aspose.Cells how to implement a custom calculation engine with result caching in C# | C# example of caching user‑defined function results in Aspose.Cells workbook | Avoid redundant formula calculations using a custom engine in Aspose.Cells | Cache identical sub‑expression evaluations for custom functions in Aspose.Cells | Performance optimization for Aspose.Cells custom functions by reusing previous results
// Tags: Aspose.Cells custom calculation engine caching | C# dictionary cache for user-defined functions | reuse formula results Aspose.Cells | cache key generation based on function arguments | performance optimization Aspose.Cells custom functions

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example defines a CachingEngine class that inherits AbstractCalculationEngine, builds a unique cache key from the function name and its parameters, returns a cached sum when the same argument set is encountered, and stores new results otherwise. The engine is attached via CalculationOptions and used with Workbook.CalculateFormula, showing cache hits for repeated MYCACHEFUNC calls and improving calculation performance.
class CachingEngine : AbstractCalculationEngine
{
    // Cache keyed by a string that represents the function name and its parameter values
    private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

    public override void Calculate(CalculationData data)
    {
        // Handle only the custom function MYCACHEFUNC
        if (data.FunctionName.Equals("MYCACHEFUNC", StringComparison.OrdinalIgnoreCase))
        {
            // Build a unique cache key
            var keyParts = new List<string> { data.FunctionName };
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is ReferredArea area)
                {
                    // Assume single‑cell area for simplicity
                    object val = area.GetValue(0, 0);
                    keyParts.Add(Convert.ToString(val));
                }
                else
                {
                    keyParts.Add(Convert.ToString(param));
                }
            }
            string cacheKey = string.Join("|", keyParts);

            // Return cached result if it exists
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // Compute the result (sum of all parameters)
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object p = data.GetParamValue(i);
                double d;
                if (p is ReferredArea ra)
                {
                    d = Convert.ToDouble(ra.GetValue(0, 0));
                }
                else
                {
                    d = Convert.ToDouble(p);
                }
                sum += d;
            }

            // Store and return the result
            data.CalculatedValue = sum;
            _cache[cacheKey] = sum;
        }
    }

    // No need to force recalculation; cache will be used when possible
    public override bool ForceRecalculate(string functionName) => false;
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate cells with values
        ws.Cells["A1"].PutValue(10);
        ws.Cells["A2"].PutValue(20);
        ws.Cells["A3"].PutValue(10);
        ws.Cells["A4"].PutValue(20);

        // Use the custom function in several cells; some have identical sub‑expressions
        ws.Cells["B1"].Formula = "=MYCACHEFUNC(A1, A2)"; // first evaluation
        ws.Cells["B2"].Formula = "=MYCACHEFUNC(A3, A4)"; // identical values, should hit cache
        ws.Cells["B3"].Formula = "=MYCACHEFUNC(A1, A2)"; // same as B1, cache again

        // Set calculation options with the caching engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new CachingEngine(),
            IgnoreError = false,
            Recursive = true
        };

        // Calculate all formulas in the workbook
        wb.CalculateFormula(opts);

        // Display the results
        Console.WriteLine("B1 = " + ws.Cells["B1"].Value);
        Console.WriteLine("B2 = " + ws.Cells["B2"].Value);
        Console.WriteLine("B3 = " + ws.Cells["B3"].Value);

        // Save the workbook (uses the provided lifecycle rule)
        wb.Save("CachingEngineDemo.xlsx");
    }
}
