using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace MaxIfsComparison
{
    // Custom calculation engine that processes built‑in functions and caches MAXIFS results
    public class CachingEngine : AbstractCalculationEngine
    {
        private readonly bool _processBuiltIn;
        // Simple in‑memory cache: key -> result
        private static readonly Dictionary<string, double> _cache = new Dictionary<string, double>();

        public CachingEngine(bool processBuiltIn)
        {
            _processBuiltIn = processBuiltIn;
        }

        public override bool ProcessBuiltInFunctions => _processBuiltIn;

        public override void Calculate(CalculationData data)
        {
            // Handle only the MAXIFS function; let the default engine process everything else
            if (!data.FunctionName.Equals("MAXIFS", StringComparison.OrdinalIgnoreCase))
                return;

            // Build a cache key from all parameters
            string key = BuildCacheKey(data);
            if (_cache.TryGetValue(key, out double cachedResult))
            {
                data.CalculatedValue = cachedResult;
                return;
            }

            // Parameter layout (simplified):
            // 0 – range to evaluate (ReferredArea)
            // 1 – criteria_range (ReferredArea)
            // 2 – criteria (string or numeric)
            if (data.ParamCount < 3)
            {
                data.CalculatedValue = "#VALUE!";
                return;
            }

            // Get the areas
            ReferredArea evalArea = (ReferredArea)data.GetParamValue(0);
            ReferredArea criteriaArea = (ReferredArea)data.GetParamValue(1);
            object criteriaObj = data.GetParamValue(2);
            string criteria = criteriaObj?.ToString() ?? string.Empty;

            double max = double.MinValue;
            bool anyMatch = false;

            // Iterate over the intersecting cells of evalArea and criteriaArea
            for (int r = 0; r <= evalArea.EndRow - evalArea.StartRow; r++)
            {
                for (int c = 0; c <= evalArea.EndColumn - evalArea.StartColumn; c++)
                {
                    object critValObj = criteriaArea.GetValue(r, c);
                    if (critValObj == null) continue;

                    if (CriteriaMatches(critValObj, criteria))
                    {
                        object evalValObj = evalArea.GetValue(r, c);
                        if (evalValObj == null) continue;

                        if (double.TryParse(evalValObj.ToString(), out double evalVal))
                        {
                            anyMatch = true;
                            if (evalVal > max) max = evalVal;
                        }
                    }
                }
            }

            data.CalculatedValue = anyMatch ? (object)max : (object)double.NaN;
            // Store in cache for future calls with the same parameters
            _cache[key] = anyMatch ? max : double.NaN;
        }

        // Simple criteria evaluator supporting >, >=, <, <=, =, <> operators
        private bool CriteriaMatches(object cellValue, string criteria)
        {
            if (cellValue == null) return false;

            // Trim whitespace
            criteria = criteria.Trim();

            // Numeric comparison
            if (double.TryParse(cellValue.ToString(), out double cellNumber) &&
                double.TryParse(criteria.TrimStart('>', '<', '=', '!'), out double critNumber))
            {
                if (criteria.StartsWith(">=")) return cellNumber >= critNumber;
                if (criteria.StartsWith("<=")) return cellNumber <= critNumber;
                if (criteria.StartsWith("<>")) return cellNumber != critNumber;
                if (criteria.StartsWith(">"))  return cellNumber > critNumber;
                if (criteria.StartsWith("<"))  return cellNumber < critNumber;
                if (criteria.StartsWith("="))  return Math.Abs(cellNumber - critNumber) < 1e-9;
                // If no operator, treat as equality
                return Math.Abs(cellNumber - critNumber) < 1e-9;
            }

            // String comparison (only equality supported)
            if (criteria.StartsWith("="))
                return cellValue.ToString().Equals(criteria.Substring(1), StringComparison.OrdinalIgnoreCase);
            return cellValue.ToString().Equals(criteria, StringComparison.OrdinalIgnoreCase);
        }

        private string BuildCacheKey(CalculationData data)
        {
            // Concatenate the address of each ReferredArea and the criteria text
            var parts = new List<string>();
            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);
                if (param is ReferredArea ra)
                {
                    parts.Add($"{ra.StartRow}:{ra.StartColumn}-{ra.EndRow}:{ra.EndColumn}");
                }
                else
                {
                    parts.Add(param?.ToString() ?? "null");
                }
            }
            return string.Join("|", parts);
        }

        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook and populate data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Values to evaluate
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);
            ws.Cells["A4"].PutValue(40);
            ws.Cells["A5"].PutValue(50);

            // Corresponding criteria values
            ws.Cells["B1"].PutValue(1);
            ws.Cells["B2"].PutValue(4);
            ws.Cells["B3"].PutValue(5);
            ws.Cells["B4"].PutValue(2);
            ws.Cells["B5"].PutValue(6);

            // Formula using MAXIFS: max of A where B > 3
            ws.Cells["C1"].Formula = "=MAXIFS(A1:A5,B1:B5,\">3\")";

            // ---------- Default engine calculation ----------
            wb.CalculateFormula(); // uses built‑in engine
            object defaultResult = ws.Cells["C1"].Value;
            Console.WriteLine($"Default engine MAXIFS result: {defaultResult}");

            // ---------- Custom engine with caching ----------
            // Enable access cache for formula calculation to improve performance
            wb.StartAccessCache(AccessCacheOptions.CalculateFormula);

            var customEngine = new CachingEngine(processBuiltIn: true);
            var calcOptions = new CalculationOptions { CustomEngine = customEngine };

            // Re‑calculate using the custom engine
            wb.CalculateFormula(calcOptions);
            object customResult = ws.Cells["C1"].Value;
            Console.WriteLine($"Custom caching engine MAXIFS result: {customResult}");

            // Close the access cache
            wb.CloseAccessCache(AccessCacheOptions.CalculateFormula);

            // Compare results
            if (defaultResult is double d1 && customResult is double d2)
            {
                double diff = Math.Abs(d1 - d2);
                Console.WriteLine($"Difference between engines: {diff}");
            }
            else
            {
                Console.WriteLine("One of the results is not a numeric value.");
            }

            // Save the workbook (lifecycle rule)
            wb.Save("MaxIfsComparison.xlsx");
        }
    }
}