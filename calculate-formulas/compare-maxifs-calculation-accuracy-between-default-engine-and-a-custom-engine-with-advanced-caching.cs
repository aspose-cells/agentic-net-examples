// Title: Aspose.Cells C# Demo: Compare Default and Custom Cached MAXIFS Calculation Engines
// Description: Creates a workbook, fills numeric data, applies a MAXIFS formula, calculates it with Aspose.Cells' built‑in engine, then re‑calculates using a custom MaxIfsEngine that caches identical calls. The sample prints both results to confirm accuracy and saves the workbook.
// Keywords: Aspose.Cells | C# | MAXIFS | custom calculation engine | caching | AbstractCalculationEngine | AccessCacheOptions | formula performance | Excel MAXIFS example | compare default engine
// Common Searches: Aspose.Cells custom MAXIFS engine C# | how to cache formula calculations in Aspose.Cells | default vs custom calculation engine Aspose.Cells | MAXIFS performance benchmark .NET | use AccessCacheOptions with custom engine Aspose.Cells
// Developer Intent: Validate that a custom cached MAXIFS engine returns the same value as the built‑in engine and assess its calculation correctness and performance.
// Use Cases: Run side‑by‑side calculations to ensure result parity between default and custom engines. | Leverage AccessCacheOptions.CalculateFormula to accelerate repeated MAXIFS evaluations. | Extend MaxIfsEngine to handle multiple criteria ranges while preserving cache logic.
// AI Prompts: Generate a C# unit test that asserts the default and custom MAXIFS results are identical for the sample worksheet. | Refactor MaxIfsEngine to support an arbitrary number of criteria pairs without breaking the existing cache mechanism. | Explain how AccessCacheOptions.CalculateFormula interacts with a custom AbstractCalculationEngine in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace MaxIfsComparisonDemo
{
    // Custom calculation engine that processes MAXIFS with a simple cache
    // Creates a workbook, fills numeric data, applies a MAXIFS formula, calculates it with Aspose.Cells' built‑in engine, then re‑calculates using a custom MaxIfsEngine that caches identical calls. The sample prints both results to confirm accuracy and saves the workbook.
    public class MaxIfsEngine : AbstractCalculationEngine
    {
        // Enable processing of built‑in functions
        public override bool ProcessBuiltInFunctions => true;

        // Simple cache: key = concatenated parameter representations, value = calculated result
        private readonly Dictionary<string, double> _cache = new Dictionary<string, double>();

        public override void Calculate(CalculationData data)
        {
            // Handle only the MAXIFS function; other functions fall back to the default engine
            if (!data.FunctionName.Equals("MAXIFS", StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key from all parameters
            var keyParts = new List<string>();
            for (int i = 0; i < data.ParamCount; i++)
            {
                // For range parameters we use their address; for literals we use their text
                object param = data.GetParamValue(i);
                if (param is ReferredArea area)
                {
                    keyParts.Add($"{area.StartRow}:{area.StartColumn}-{area.EndRow}:{area.EndColumn}");
                }
                else
                {
                    keyParts.Add(data.GetParamText(i));
                }
            }
            string cacheKey = string.Join("|", keyParts);

            // Return cached result if present
            if (_cache.TryGetValue(cacheKey, out double cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // MAXIFS syntax: MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2] …)
            // For demo purposes we support only one criteria pair.
            if (data.ParamCount < 3)
            {
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // First parameter: range to evaluate for maximum
            ReferredArea maxRange = (ReferredArea)data.GetParamValue(0);
            // Second parameter: criteria range
            ReferredArea criteriaRange = (ReferredArea)data.GetParamValue(1);
            // Third parameter: criteria (string or number)
            string criteria = data.GetParamText(2).Trim('\"');

            double maxValue = double.MinValue;
            bool anyMatch = false;

            // Iterate over the cells of the ranges (they must be of the same size)
            for (int r = maxRange.StartRow; r <= maxRange.EndRow; r++)
            {
                for (int c = maxRange.StartColumn; c <= maxRange.EndColumn; c++)
                {
                    object critCellObj = criteriaRange.GetValue(r, c);
                    double critCellVal = Convert.ToDouble(critCellObj);

                    // Simple criteria handling: only supports ">=N", "<=N", ">N", "<N", "=N"
                    bool match = false;
                    if (criteria.StartsWith(">="))
                        match = critCellVal >= Convert.ToDouble(criteria.Substring(2));
                    else if (criteria.StartsWith("<="))
                        match = critCellVal <= Convert.ToDouble(criteria.Substring(2));
                    else if (criteria.StartsWith(">"))
                        match = critCellVal > Convert.ToDouble(criteria.Substring(1));
                    else if (criteria.StartsWith("<"))
                        match = critCellVal < Convert.ToDouble(criteria.Substring(1));
                    else if (criteria.StartsWith("="))
                        match = critCellVal == Convert.ToDouble(criteria.Substring(1));
                    else
                        match = critCellObj.ToString() == criteria; // fallback to string equality

                    if (match)
                    {
                        anyMatch = true;
                        object maxCellObj = maxRange.GetValue(r, c);
                        double maxCellVal = Convert.ToDouble(maxCellObj);
                        if (maxCellVal > maxValue)
                            maxValue = maxCellVal;
                    }
                }
            }

            data.CalculatedValue = anyMatch ? (object)maxValue : "#N/A";

            // Store result in cache for future identical calls
            _cache[cacheKey] = anyMatch ? maxValue : double.NaN;
        }

        // No special forced recalculation logic needed
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and populate data --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Values for MAXIFS range (A1:A6)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);
            cells["A5"].PutValue(50);
            cells["A6"].PutValue(60);

            // Corresponding criteria range (B1:B6)
            cells["B1"].PutValue(1);
            cells["B2"].PutValue(3);
            cells["B3"].PutValue(5);
            cells["B4"].PutValue(3);
            cells["B5"].PutValue(7);
            cells["B6"].PutValue(3);

            // Formula using MAXIFS: find max in A where B >= 3
            string maxIfsFormula = "=MAXIFS(A1:A6,B1:B6,\">=3\")";
            cells["C1"].Formula = maxIfsFormula;

            // -------------------- Default engine calculation --------------------
            workbook.CalculateFormula(); // default calculation
            object defaultResult = cells["C1"].Value;
            Console.WriteLine($"Default engine result: {defaultResult}");

            // -------------------- Custom engine with caching --------------------
            // Start access cache for calculation to improve performance
            workbook.StartAccessCache(AccessCacheOptions.CalculateFormula);

            CalculationOptions customOptions = new CalculationOptions
            {
                CustomEngine = new MaxIfsEngine()
            };

            // Re‑calculate using the custom engine
            workbook.CalculateFormula(customOptions);
            object customResult = cells["C1"].Value;
            Console.WriteLine($"Custom engine result: {customResult}");

            // Close the access cache
            workbook.CloseAccessCache(AccessCacheOptions.CalculateFormula);

            // -------------------- Save workbook --------------------
            workbook.Save("MaxIfsComparisonResult.xlsx");
        }
    }
}
