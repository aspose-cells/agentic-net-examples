// Title: Cache Results of a User‑Defined Formula in Aspose.Cells with C#
// Description: The sample extends Aspose.Cells' AbstractCalculationEngine to memoize the output of a user‑defined formula. A deterministic key derived from all arguments—including range values—is used to store the result in a dictionary. Repeated invocations with identical parameters retrieve the cached value instantly, while changes to unrelated cells do not force recomputation. The workbook demonstrates the behavior with the MYCACHEDFUNC function and a compute‑counter.
// Keywords: Aspose.Cells | C# memoization | calculation engine | dictionary cache | .NET spreadsheet performance | custom function | ReferredArea serialization | ForceRecalculate | custom engine
// Common Searches: Aspose.Cells memoize custom formula | C# calculation engine caching example | avoid duplicate evaluation in spreadsheet .NET | how to store function results in Aspose.Cells | performance boost for user‑defined functions Aspose.Cells
// Developer Intent: Create a caching layer for a user‑defined spreadsheet function in Aspose.Cells to reduce redundant processing.
// Use Cases: Accelerate workbooks where the same user‑defined formula appears in many cells with the same arguments. | Validate cache effectiveness by monitoring a computation counter before and after recalculation. | Handle range inputs by converting cell contents into a unique identifier for the cache. | Maintain correct results when only non‑dependent cells are modified.
// AI Prompts: Generate MSTest code that confirms cache hits and that the computation counter does not increase on a second workbook calculation. | Describe how to extend the key builder to include cell addresses for more precise memoization. | Explain how to adjust ForceRecalculate so that only volatile functions bypass the cache. | Provide instructions for adding the CachedEngine to an existing Aspose.Cells solution.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace CustomFunctionCachingDemo
{
    // Custom calculation engine that caches results of a user‑defined function.
    // The sample extends Aspose.Cells' AbstractCalculationEngine to memoize the output of a user‑defined formula. A deterministic key derived from all arguments—including range values—is used to store the result in a dictionary. Repeated invocations with identical parameters retrieve the cached value instantly, while changes to unrelated cells do not force recomputation. The workbook demonstrates the behavior with the MYCACHEDFUNC function and a compute‑counter.
    public class CachedEngine : AbstractCalculationEngine
    {
        // Simple cache: key = concatenated parameter values, value = calculated result.
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>();

        // Counter to demonstrate how many times the function is actually computed.
        private int _computeCount = 0;

        // The custom function name we want to cache.
        private const string FunctionName = "MYCACHEDFUNC";

        // Do not force recalculation for this function – allow caching.
        public override bool ForceRecalculate(string functionName)
        {
            return false; // return true only for volatile functions.
        }

        public override void Calculate(CalculationData data)
        {
            // Only handle our custom function.
            if (!string.Equals(data.FunctionName, FunctionName, StringComparison.OrdinalIgnoreCase))
                return; // let the default engine handle other functions.

            // Build a cache key from the function parameters.
            string key = BuildCacheKey(data);

            // If we have a cached value, return it.
            if (_cache.TryGetValue(key, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // ----- Actual calculation (executed only when cache miss) -----
            _computeCount++;

            // Example calculation: sum of all numeric parameters.
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);

                // Parameters may be scalar values or ReferredArea objects.
                if (param is ReferredArea area)
                {
                    // For simplicity, take the value at the top‑left cell of the area.
                    object val = area.GetValue(0, 0);
                    if (val is double d)
                        sum += d;
                    else if (double.TryParse(val?.ToString(), out d))
                        sum += d;
                }
                else if (param is double d)
                {
                    sum += d;
                }
                else if (double.TryParse(param?.ToString(), out double d2))
                {
                    sum += d2;
                }
            }

            // Store the result in the cache and set it as the function result.
            _cache[key] = sum;
            data.CalculatedValue = sum;
        }

        // Helper to create a deterministic string key from all parameters.
        private string BuildCacheKey(CalculationData data)
        {
            var parts = new List<string>();
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is ReferredArea area)
                {
                    // Serialize the whole area (row‑major) to capture its content.
                    for (int r = 0; r <= area.EndRow - area.StartRow; r++)
                    {
                        for (int c = 0; c <= area.EndColumn - area.StartColumn; c++)
                        {
                            object val = area.GetValue(r, c);
                            parts.Add(val?.ToString() ?? "null");
                        }
                    }
                }
                else
                {
                    parts.Add(param?.ToString() ?? "null");
                }
            }
            return string.Join("|", parts);
        }

        // Expose the compute count for demonstration purposes.
        public int ComputeCount => _computeCount;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and fill some data.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["B1"].PutValue(10);
            ws.Cells["B2"].PutValue(20);
            ws.Cells["B3"].PutValue(30);

            // Use the custom function in several cells with identical inputs.
            ws.Cells["A1"].Formula = "=MYCACHEDFUNC(B1)";
            ws.Cells["A2"].Formula = "=MYCACHEDFUNC(B1)";
            ws.Cells["A3"].Formula = "=MYCACHEDFUNC(B2)";
            ws.Cells["A4"].Formula = "=MYCACHEDFUNC(B2)";
            ws.Cells["A5"].Formula = "=MYCACHEDFUNC(B3)";

            // Set up calculation options with our cached engine.
            var engine = new CachedEngine();
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = engine
            };

            // First calculation – cache will be populated.
            wb.CalculateFormula(options);
            Console.WriteLine("After first calculation:");
            PrintResults(ws);
            Console.WriteLine($"Engine performed actual computation {engine.ComputeCount} time(s).");

            // Change a cell that is NOT used by the formulas to prove cache reuse.
            ws.Cells["C1"].PutValue(999);

            // Second calculation – cached results should be reused, compute count unchanged.
            wb.CalculateFormula(options);
            Console.WriteLine("\nAfter second calculation (no relevant data changed):");
            PrintResults(ws);
            Console.WriteLine($"Engine performed actual computation {engine.ComputeCount} time(s).");

            // Save the workbook.
            wb.Save("CachedFunctionDemo.xlsx");
        }

        // Helper to display the values of the cells that use the custom function.
        private static void PrintResults(Worksheet ws)
        {
            for (int row = 0; row < 5; row++)
            {
                Console.WriteLine($"A{row + 1} = {ws.Cells[row, 0].Value}");
            }
        }
    }
}
