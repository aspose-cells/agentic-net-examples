using System;
using System.Collections.Concurrent;
using Aspose.Cells;

namespace AsposeCellsOptimizationDemo
{
    // Custom engine that implements a fast, cache‑aware function.
    public class FastEngine : AbstractCalculationEngine
    {
        // Cache key: function name + serialized parameters.
        private static readonly ConcurrentDictionary<string, object> _cache = new();

        // The custom function is not volatile – it can be cached for shared formulas.
        public override bool ForceRecalculate(string functionName) => false;

        // Do not process built‑in functions – keep default engine for performance.
        public override bool ProcessBuiltInFunctions => false;

        public override void Calculate(CalculationData data)
        {
            // Handle only our custom function.
            if (!string.Equals(data.FunctionName, "MYFASTFUNC", StringComparison.OrdinalIgnoreCase))
                return; // Let the default engine handle other functions.

            // Build a cache key from the parameter values.
            // Using GetParamValue ensures parameters are already evaluated.
            var keyBuilder = new System.Text.StringBuilder("MYFASTFUNC");
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                // Simple conversion to string – adjust if you need higher precision.
                keyBuilder.Append('|').Append(param?.ToString() ?? "null");
            }
            string cacheKey = keyBuilder.ToString();

            // Try to get a cached result.
            if (_cache.TryGetValue(cacheKey, out object cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // ----- Expensive calculation placeholder -----
            // Replace this block with the real heavy logic.
            // For demonstration we just sum numeric parameters.
            double sum = 0;
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is double d)
                    sum += d;
                else if (double.TryParse(param?.ToString(), out double parsed))
                    sum += parsed;
            }
            // -------------------------------------------

            // Store result in cache and assign to the cell.
            _cache[cacheKey] = sum;
            data.CalculatedValue = sum;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a workbook and fill data.
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate 10,000 rows with simple values.
            for (int r = 0; r < 10000; r++)
                cells[r, 0].PutValue(r + 1); // Column A

            // Add a shared formula that uses the custom function.
            // The formula will be the same for all rows, enabling caching.
            string sharedFormula = "=MYFASTFUNC(A{0})";
            cells[0, 1].SetSharedFormula(string.Format(sharedFormula, 1), 10000, 1); // Column B

            // -------------------------------------------------
            // 2. Configure calculation options with the custom engine.
            // -------------------------------------------------
            CalculationOptions opts = new CalculationOptions
            {
                CustomEngine = new FastEngine(),
                // Larger stack size reduces recursion checks for deep dependency trees.
                CalcStackSize = 500,
                // Keep default values for other options (IgnoreError = true, Recursive = true).
            };

            // -------------------------------------------------
            // 3. Perform calculation.
            // -------------------------------------------------
            // Using Workbook.CalculateFormula applies the engine to all cells.
            wb.CalculateFormula(opts);

            // -------------------------------------------------
            // 4. Verify a few results (optional).
            // -------------------------------------------------
            Console.WriteLine("B1 = " + cells["B1"].Value); // Should be 1
            Console.WriteLine("B5000 = " + cells[4999, 1].Value); // Should be 5000
            Console.WriteLine("B10000 = " + cells[9999, 1].Value); // Should be 10000

            // -------------------------------------------------
            // 5. Save the workbook (lifecycle rule compliance).
            // -------------------------------------------------
            wb.Save("OptimizedCalculation.xlsx");
        }
    }
}