// Title: Compare default and custom caching calculation engines for MAXIFS in Aspose.Cells .NET
// AI Prompts: Create a workbook, populate sample data, apply =MAXIFS(A1:A3,B1:B3,2) and calculate it with the built‑in engine, then recalculate using a user‑defined calculation engine that caches results and display both values. | Implement a subclass of AbstractCalculationEngine that intercepts MAXIFS, stores computed maxima in a dictionary cache, configure CalculationOptions.CustomEngine, and invoke workbook.CalculateFormula with these options. | Extend the user‑defined MAXIFS engine to support multiple criteria ranges, run the formula, and verify that the cached output matches the built‑in engine's result.
// Common Searches: Aspose.Cells how to use a user‑defined calculation engine for MAXIFS | compare default formula calculation with custom engine in Aspose.Cells .NET | cache MAXIFS results using AccessCacheOptions in Aspose.Cells | override built‑in functions in Aspose.Cells calculation engine example | performance test for MAXIFS with custom caching engine Aspose.Cells
// Tags: MAXIFS calculation engine with caching Aspose.Cells | AccessCacheOptions formula caching .NET | default vs custom calculation engine Aspose.Cells | MAXIFS performance optimization Aspose.Cells | override built‑in functions Aspose.Cells | Aspose.Cells calculation options custom engine

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting sample data, evaluating a MAXIFS formula with the default engine, then recalculating the same formula using a custom caching engine, comparing the outcomes, and saving the result.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for MAXIFS
        // Values to evaluate
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        // Criteria range
        sheet.Cells["B1"].PutValue(1);
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["B3"].PutValue(3);

        // Formula: return the maximum value in A1:A3 where corresponding B cell equals 2 (expected 20)
        string maxIfsFormula = "=MAXIFS(A1:A3,B1:B3,2)";

        // ---------- Default calculation engine ----------
        sheet.Cells["C1"].Formula = maxIfsFormula;
        workbook.CalculateFormula(); // uses default engine
        object defaultResult = sheet.Cells["C1"].Value;

        // ---------- Custom calculation engine with advanced caching ----------
        // Start access cache for calculation to improve performance
        workbook.StartAccessCache(AccessCacheOptions.CalculateFormula);

        // Set up calculation options with the custom engine
        CalculationOptions customOptions = new CalculationOptions
        {
            CustomEngine = new CustomMaxIfsEngine()
        };

        // Place the same formula in another cell and calculate with custom engine
        sheet.Cells["D1"].Formula = maxIfsFormula;
        workbook.CalculateFormula(customOptions);
        object customResult = sheet.Cells["D1"].Value;

        // Close the access cache
        workbook.CloseAccessCache(AccessCacheOptions.CalculateFormula);

        // Output comparison
        Console.WriteLine($"Default MAXIFS result: {defaultResult}");
        Console.WriteLine($"Custom MAXIFS result: {customResult}");
        Console.WriteLine($"Results are {(Equals(defaultResult, customResult) ? "identical" : "different")}.");

        // Save the workbook (optional)
        workbook.Save("MaxIfsComparison.xlsx");
    }

    // Custom calculation engine that intercepts MAXIFS and caches results
    class CustomMaxIfsEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions so this engine receives MAXIFS calls
        public override bool ProcessBuiltInFunctions => true;

        // Simple cache: key = criteria description, value = computed max
        private readonly Dictionary<string, double> _cache = new Dictionary<string, double>();

        public override void Calculate(CalculationData data)
        {
            // Only handle MAXIFS; let the default engine process other functions
            if (!data.FunctionName.Equals("MAXIFS", StringComparison.OrdinalIgnoreCase))
                return;

            // Expect at least three parameters: range, criteria_range1, criteria1
            if (data.ParamCount < 3)
            {
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // Retrieve the main range and the first criteria range/value
            var valueRange = (ReferredArea)data.GetParamValue(0);
            var criteriaRange = (ReferredArea)data.GetParamValue(1);
            var criteria = data.GetParamValue(2);

            // Build a cache key based on the criteria range address and the criteria value
            string cacheKey = $"{criteriaRange.StartRow}:{criteriaRange.StartColumn}-{criteriaRange.EndRow}:{criteriaRange.EndColumn}:{criteria}";

            // Return cached result if available
            if (_cache.TryGetValue(cacheKey, out double cachedMax))
            {
                data.CalculatedValue = cachedMax;
                return;
            }

            double max = double.MinValue;

            // Iterate over the value range and evaluate the criteria
            for (int r = valueRange.StartRow; r <= valueRange.EndRow; r++)
            {
                for (int c = valueRange.StartColumn; c <= valueRange.EndColumn; c++)
                {
                    // Determine corresponding cell in the criteria range
                    int critRowOffset = r - valueRange.StartRow;
                    int critColOffset = c - valueRange.StartColumn;
                    object critCellValue = criteriaRange.GetValue(critRowOffset, critColOffset);

                    // Simple equality comparison for the demo (Excel supports many operators)
                    if (object.Equals(critCellValue, criteria))
                    {
                        object val = valueRange.GetValue(critRowOffset, critColOffset);
                        if (val is double d)
                        {
                            if (d > max) max = d;
                        }
                        else if (double.TryParse(val?.ToString(), out double parsed))
                        {
                            if (parsed > max) max = parsed;
                        }
                    }
                }
            }

            // If no matching cells, Excel returns 0
            if (max == double.MinValue) max = 0;

            // Store the computed max in the cache
            _cache[cacheKey] = max;

            // Set the calculated value for the formula
            data.CalculatedValue = max;
        }

        // No need to force recalculation for this function
        public override bool ForceRecalculate(string functionName) => false;
    }
}
