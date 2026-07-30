// Title: Compare Aspose.Cells MAXIFS Results: Default vs Custom Caching Engine (C#)
// Description: Creates a workbook with two numeric columns, applies a MAXIFS formula (A where B > 2), calculates it first with Aspose.Cells' built‑in engine, then with a custom AbstractCalculationEngine that caches results, and finally checks that both outcomes match before saving the file.
// Keywords: Aspose.Cells | MAXIFS | custom calculation engine | C# | formula caching | AbstractCalculationEngine | performance comparison | AccessCache | Excel formula evaluation
// Common Searches: Aspose.Cells custom MAXIFS engine example | compare default and custom formula calculation in .NET | how to cache MAXIFS results with Aspose.Cells | implement AbstractCalculationEngine for Excel functions | use AccessCache with custom calculation engine
// Developer Intent: Confirm that a user‑defined MAXIFS engine with in‑memory caching returns the same value as the library's native engine.
// Use Cases: Benchmark performance gains from caching complex formulas on large sheets | Extend MAXIFS logic (e.g., additional operators) while preserving existing results | Standardize calculation behavior across multiple workbooks in automated reporting
// AI Prompts: Write a C# unit test that validates MaxIfsCustomEngine produces identical results to the default engine for varied data ranges and criteria. | Provide a thread‑safe version of MaxIfsCustomEngine that supports multiple criteria pairs using ConcurrentDictionary for caching. | Explain how to integrate a custom calculation engine into a high‑throughput Aspose.Cells pipeline, covering cache lifecycle, error handling, and parallel processing.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace MaxIfsComparisonDemo
{
    // Custom calculation engine that processes the built‑in MAXIFS function
    // and uses a simple in‑memory cache to avoid repeated calculations.
    // Creates a workbook with two numeric columns, applies a MAXIFS formula (A where B > 2), calculates it first with Aspose.Cells' built‑in engine, then with a custom AbstractCalculationEngine that caches results, and finally checks that both outcomes match before saving the file.
    public class MaxIfsCustomEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions.
        public override bool ProcessBuiltInFunctions => true;

        // Cache key: concatenated string of range addresses and criteria.
        private static readonly Dictionary<string, double> _cache = new Dictionary<string, double>();

        public override void Calculate(CalculationData data)
        {
            // Only handle MAXIFS; other functions fall back to the default engine.
            if (!data.FunctionName.Equals("MAXIFS", StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key from all parameters.
            var keyParts = new List<string>();
            for (int i = 0; i < data.ParamCount; i++)
            {
                // For range parameters we use the address; for criteria we use the literal text.
                if (data.GetParamValue(i) is ReferredArea area)
                {
                    keyParts.Add(area.StartRow + ":" + area.StartColumn + "-" + area.EndRow + ":" + area.EndColumn);
                }
                else
                {
                    keyParts.Add(data.GetParamText(i));
                }
            }
            string cacheKey = string.Join("|", keyParts);

            // Return cached result if present.
            if (_cache.TryGetValue(cacheKey, out double cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // MAXIFS syntax: MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2]…)
            // For simplicity we support only one criteria pair.
            if (data.ParamCount < 3)
            {
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // Parameter 0: range to evaluate for maximum.
            ReferredArea maxRange = (ReferredArea)data.GetParamValue(0);
            // Parameter 1: criteria range.
            ReferredArea criteriaRange = (ReferredArea)data.GetParamValue(1);
            // Parameter 2: criteria string (e.g., ">2").
            string criteria = data.GetParamText(2).Trim('\"');

            double maxValue = double.MinValue;
            bool anyMatch = false;

            // Iterate over the cells of the criteria range; assume same size as maxRange.
            for (int r = criteriaRange.StartRow; r <= criteriaRange.EndRow; r++)
            {
                for (int c = criteriaRange.StartColumn; c <= criteriaRange.EndColumn; c++)
                {
                    object critObj = criteriaRange.GetValue(r, c);
                    double critVal = Convert.ToDouble(critObj);

                    // Evaluate the criteria expression.
                    bool meets = EvaluateCriteria(critVal, criteria);
                    if (meets)
                    {
                        anyMatch = true;
                        // Corresponding cell in maxRange.
                        int offsetRow = r - criteriaRange.StartRow;
                        int offsetCol = c - criteriaRange.StartColumn;
                        int targetRow = maxRange.StartRow + offsetRow;
                        int targetCol = maxRange.StartColumn + offsetCol;
                        object maxObj = maxRange.GetValue(targetRow, targetCol);
                        double maxVal = Convert.ToDouble(maxObj);
                        if (maxVal > maxValue)
                            maxValue = maxVal;
                    }
                }
            }

            data.CalculatedValue = anyMatch ? (object)maxValue : "#N/A";

            // Store result in cache.
            _cache[cacheKey] = anyMatch ? maxValue : double.NaN;
        }

        // Very simple criteria evaluator supporting >, >=, <, <=, = and <> operators.
        private bool EvaluateCriteria(double value, string criteria)
        {
            if (string.IsNullOrEmpty(criteria))
                return false;

            if (criteria.StartsWith(">="))
                return value >= double.Parse(criteria.Substring(2));
            if (criteria.StartsWith("<="))
                return value <= double.Parse(criteria.Substring(2));
            if (criteria.StartsWith("<>"))
                return value != double.Parse(criteria.Substring(2));
            if (criteria.StartsWith(">"))
                return value > double.Parse(criteria.Substring(1));
            if (criteria.StartsWith("<"))
                return value < double.Parse(criteria.Substring(1));
            if (criteria.StartsWith("="))
                return value == double.Parse(criteria.Substring(1));

            // If no operator, treat as equality.
            return value == double.Parse(criteria);
        }

        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create workbook and populate data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Max range (A1:A5)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(5);
            cells["A4"].PutValue(30);
            cells["A5"].PutValue(25);

            // Criteria range (B1:B5)
            cells["B1"].PutValue(1);
            cells["B2"].PutValue(3);
            cells["B3"].PutValue(2);
            cells["B4"].PutValue(4);
            cells["B5"].PutValue(5);

            // Formula cell using MAXIFS: find max in A where corresponding B > 2
            cells["C1"].Formula = "=MAXIFS(A1:A5, B1:B5, \">2\")";

            // ---------- Default engine calculation ----------
            workbook.CalculateFormula(); // uses built‑in engine
            object defaultResult = cells["C1"].Value;
            Console.WriteLine($"Default engine MAXIFS result: {defaultResult}");

            // ---------- Custom engine with caching ----------
            // Start access cache for formula calculation to improve performance
            workbook.StartAccessCache(AccessCacheOptions.CalculateFormula);

            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new MaxIfsCustomEngine(),
                IgnoreError = true,
                Recursive = true
            };

            // Re‑calculate using the custom engine
            workbook.CalculateFormula(options);
            object customResult = cells["C1"].Value;
            Console.WriteLine($"Custom engine MAXIFS result: {customResult}");

            // Close the access cache
            workbook.CloseAccessCache(AccessCacheOptions.CalculateFormula);

            // ---------- Comparison ----------
            bool areEqual = Equals(defaultResult, customResult);
            Console.WriteLine($"Results are equal: {areEqual}");

            // Save workbook (lifecycle rule)
            workbook.Save("MaxIfsComparisonResult.xlsx");
        }
    }
}
