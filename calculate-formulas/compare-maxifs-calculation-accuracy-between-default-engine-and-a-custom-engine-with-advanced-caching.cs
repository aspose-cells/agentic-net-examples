// Title: Evaluate and compare MAXIFS results using Aspose.Cells default calculation engine versus a custom cached engine in C#
// Common Searches: how to evaluate a MAXIFS formula using Aspose.Cells and a custom C# engine | compare results of Aspose.Cells default calculation engine with a custom MAXIFS cache | C# example for implementing a cached MAXIFS evaluator in Aspose.Cells | verify MAXIFS accuracy when using a custom caching engine in .NET | Aspose.Cells custom formula evaluation for MAXIFS with advanced caching
// Tags: MAXIFS calculation comparison Aspose.Cells | custom caching engine for Excel formulas C# | Aspose.Cells default formula calculation engine | range-based MAXIFS implementation with cache | Excel formula performance testing Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace MaxIfsComparison
{
    // Creates a workbook with sample categories and values, inserts a MAXIFS formula, calculates it using Aspose.Cells' built‑in engine, then evaluates the same formula with a custom MAXIFS engine that caches results, and finally compares the two outcomes for equality.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook with sample data and a MAXIFS formula.
                // -----------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");

                // Sample data (Category, Value)
                string[] categories = { "X", "Y", "X", "Z", "X", "Y", "Z", "X", "Y", "Z" };
                double[] values = { 10, 20, 15, 5, 25, 30, 12, 18, 22, 8 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 2, 0].PutValue(categories[i]); // Column A
                    cells[i + 2, 1].PutValue(values[i]);    // Column B
                }

                // Place the MAXIFS formula in C2
                // Formula: =MAXIFS(B2:B11, A2:A11, "X")
                string formula = "=MAXIFS(B2:B11,A2:A11,\"X\")";
                cells["C2"].Formula = formula;

                // -----------------------------------------------------------------
                // 2. Calculate using the default Aspose.Cells calculation engine.
                // -----------------------------------------------------------------
                wb.CalculateFormula();
                object defaultResult = cells["C2"].Value;
                Console.WriteLine($"Default engine result for {formula} : {defaultResult}");

                // -----------------------------------------------------------------
                // 3. Calculate using a custom engine with advanced caching.
                // -----------------------------------------------------------------
                CustomMaxIfsEngine engine = new CustomMaxIfsEngine(ws);
                object customResult = engine.EvaluateFormula(formula);
                Console.WriteLine($"Custom engine result for {formula} : {customResult}");

                // -----------------------------------------------------------------
                // 4. Compare the two results.
                // -----------------------------------------------------------------
                bool areEqual = AreValuesEqual(defaultResult, customResult);
                Console.WriteLine($"Results are {(areEqual ? "identical" : "different")}.");

                // Keep console window open when run outside of debugger.
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static bool AreValuesEqual(object a, object b, double tolerance = 1e-9)
        {
            if (a == null || b == null) return false;

            double da, db;
            try
            {
                da = Convert.ToDouble(a);
                db = Convert.ToDouble(b);
            }
            catch
            {
                // If conversion fails, fall back to string comparison.
                return a.ToString() == b.ToString();
            }

            return Math.Abs(da - db) <= tolerance;
        }
    }

    class CustomMaxIfsEngine
    {
        private readonly Worksheet _worksheet;
        // Cache key format: "valueRange|criteriaRange|criteria"
        private readonly Dictionary<string, double> _cache = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        public CustomMaxIfsEngine(Worksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        /// <summary>
        /// Evaluates a formula string that is expected to be a single MAXIFS call.
        /// If the formula is not a MAXIFS call, the method returns null.
        /// </summary>
        public object EvaluateFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return null;

            // Remove leading '=' if present.
            string trimmed = formula.TrimStart('=').Trim();

            // Quick check – does it start with MAXIFS?
            if (!trimmed.StartsWith("MAXIFS", StringComparison.OrdinalIgnoreCase))
                return null;

            // Parse arguments.
            var args = ParseMaxIfsArguments(trimmed);
            if (args == null)
                return null;

            string cacheKey = $"{args.ValueRange}|{args.CriteriaRange}|{args.Criteria}";
            if (_cache.TryGetValue(cacheKey, out double cachedResult))
            {
                return cachedResult;
            }

            double result = ComputeMaxIfs(args.ValueRange, args.CriteriaRange, args.Criteria);
            _cache[cacheKey] = result;
            return result;
        }

        /// <summary>
        /// Holds the three arguments required for MAXIFS.
        /// </summary>
        private class MaxIfsArgs
        {
            public string ValueRange { get; set; } = string.Empty;
            public string CriteriaRange { get; set; } = string.Empty;
            public string Criteria { get; set; } = string.Empty;
        }

        /// <summary>
        /// Parses a MAXIFS formula string and extracts the three arguments.
        /// Expected format: MAXIFS(value_range, criteria_range, criteria)
        /// Supports criteria that are quoted strings or numbers.
        /// </summary>
        private MaxIfsArgs ParseMaxIfsArguments(string formulaBody)
        {
            int firstParen = formulaBody.IndexOf('(');
            int lastParen = formulaBody.LastIndexOf(')');
            if (firstParen < 0 || lastParen < 0 || lastParen <= firstParen)
                return null;

            string inner = formulaBody.Substring(firstParen + 1, lastParen - firstParen - 1);
            List<string> parts = new List<string>();
            int start = 0;
            bool inQuotes = false;
            for (int i = 0; i < inner.Length; i++)
            {
                if (inner[i] == '\"')
                    inQuotes = !inQuotes;
                else if (inner[i] == ',' && !inQuotes)
                {
                    parts.Add(inner.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }
            parts.Add(inner.Substring(start).Trim());

            if (parts.Count != 3)
                return null;

            // Strip surrounding quotes from criteria if present.
            string criteria = parts[2];
            if (criteria.StartsWith("\"") && criteria.EndsWith("\"") && criteria.Length >= 2)
                criteria = criteria.Substring(1, criteria.Length - 2);

            return new MaxIfsArgs
            {
                ValueRange = parts[0],
                CriteriaRange = parts[1],
                Criteria = criteria
            };
        }

        /// <summary>
        /// Performs the actual MAXIFS calculation by scanning the value range and
        /// applying the criteria against the criteria range.
        /// </summary>
        private double ComputeMaxIfs(string valueRange, string criteriaRange, string criteria)
        {
            // Create ranges.
            AsposeRange values = _worksheet.Cells.CreateRange(valueRange);
            AsposeRange criteriaRng = _worksheet.Cells.CreateRange(criteriaRange);

            // Ensure both ranges have the same number of cells.
            if (values.RowCount != criteriaRng.RowCount || values.ColumnCount != criteriaRng.ColumnCount)
                throw new InvalidOperationException("Value range and criteria range must be of the same size.");

            double max = double.MinValue;
            bool anyMatch = false;

            for (int row = 0; row < values.RowCount; row++)
            {
                for (int col = 0; col < values.ColumnCount; col++)
                {
                    object critObj = criteriaRng[row, col].Value;
                    string critStr = critObj?.ToString() ?? string.Empty;

                    // Simple equality comparison (Excel's MAXIFS supports many operators,
                    // but for this demo we handle only plain equality).
                    if (string.Equals(critStr, criteria, StringComparison.OrdinalIgnoreCase))
                    {
                        object valObj = values[row, col].Value;
                        if (valObj != null && double.TryParse(valObj.ToString(), out double val))
                        {
                            if (val > max)
                                max = val;
                            anyMatch = true;
                        }
                    }
                }
            }

            // If no matching rows, Excel returns 0.
            return anyMatch ? max : 0.0;
        }
    }
}
