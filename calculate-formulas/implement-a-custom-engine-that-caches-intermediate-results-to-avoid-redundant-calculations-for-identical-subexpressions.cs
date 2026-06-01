using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCachingEngineDemo
{
    // Custom calculation engine that caches results of identical sub‑expressions
    public class CachingEngine : AbstractCalculationEngine
    {
        // Simple in‑memory cache: key -> calculated value
        private readonly Dictionary<string, object> _cache = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        // This engine does not need literal text or array mode for parameters
        public override bool IsParamLiteralRequired => false;
        public override bool IsParamArrayModeRequired => false;

        // We are not processing built‑in functions; only custom ones are handled here
        public override bool ProcessBuiltInFunctions => false;

        // Never force recalculation; allow cached values to be reused
        public override bool ForceRecalculate(string functionName) => false;

        public override void Calculate(CalculationData data)
        {
            // Only handle our custom function "MYFUNC"
            if (!data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let default engine handle other functions

            // Build a cache key based on function name and parameter values
            var keyParts = new List<string> { data.FunctionName.ToUpper() };
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                // Parameters may be ReferredArea (cell/range) or direct values
                if (param is ReferredArea area)
                {
                    // For simplicity, take the top‑left cell value
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

            // Perform the actual calculation (sum of two parameters)
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                double val;
                if (param is ReferredArea area)
                {
                    val = Convert.ToDouble(area.GetValue(0, 0));
                }
                else
                {
                    val = Convert.ToDouble(param);
                }
                sum += val;
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
            // Create a new workbook and populate some cells
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(10); // Same value as A1 to test caching

            // Use the custom function MYFUNC in several cells
            sheet.Cells["B1"].Formula = "=MYFUNC(A1, A2)"; // 10 + 20 = 30
            sheet.Cells["B2"].Formula = "=MYFUNC(A1, A2)"; // Same parameters -> cached
            sheet.Cells["B3"].Formula = "=MYFUNC(A3, A2)"; // 10 + 20 = 30, same result but different reference -> cached separately

            // Set calculation options to use our caching engine
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new CachingEngine(),
                IgnoreError = false,
                Recursive = true
            };

            // Calculate all formulas
            workbook.CalculateFormula(options);

            // Output results to verify caching worked (values should be 30)
            Console.WriteLine("B1 = " + sheet.Cells["B1"].Value); // 30
            Console.WriteLine("B2 = " + sheet.Cells["B2"].Value); // 30 (cached)
            Console.WriteLine("B3 = " + sheet.Cells["B3"].Value); // 30 (cached separately)

            // Save the workbook (creation and saving follow the standard lifecycle)
            workbook.Save("CachingEngineDemo.xlsx");
        }
    }
}