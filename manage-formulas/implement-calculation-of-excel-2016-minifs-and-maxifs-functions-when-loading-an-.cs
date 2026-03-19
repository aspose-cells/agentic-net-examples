using System;
using System.Collections.Generic;
using Aspose.Cells;

class MinMaxIfsEngine : AbstractCalculationEngine
{
    // Enable processing of built‑in functions so our engine is called for MINIFS/MAXIFS
    public override bool ProcessBuiltInFunctions => true;

    public override void Calculate(CalculationData data)
    {
        string fn = data.FunctionName.ToUpperInvariant();

        if (fn != "MINIFS" && fn != "MAXIFS")
            return; // let the default engine handle other functions

        // First argument: range to evaluate
        ReferredArea evalRange = (ReferredArea)data.GetParamValue(0);

        // Remaining arguments are pairs: criteria_range, criteria
        var critRanges = new List<ReferredArea>();
        var critValues = new List<object>();
        for (int i = 1; i < data.ParamCount; i += 2)
        {
            critRanges.Add((ReferredArea)data.GetParamValue(i));
            critValues.Add(data.GetParamValue(i + 1));
        }

        double result = fn == "MINIFS" ? double.MaxValue : double.MinValue;
        bool anyMatch = false;

        // Iterate over each cell in the evaluation range
        for (int r = evalRange.StartRow; r <= evalRange.EndRow; r++)
        {
            for (int c = evalRange.StartColumn; c <= evalRange.EndColumn; c++)
            {
                object cellVal = evalRange.GetValue(r, c);
                if (!IsNumeric(cellVal))
                    continue; // ignore non‑numeric cells

                bool matchesAll = true;
                // Check each criteria pair
                for (int k = 0; k < critRanges.Count; k++)
                {
                    ReferredArea critRange = critRanges[k];
                    object crit = critValues[k];
                    object critCellVal = critRange.GetValue(r, c);
                    if (!EvaluateCriteria(critCellVal, crit))
                    {
                        matchesAll = false;
                        break;
                    }
                }

                if (matchesAll)
                {
                    anyMatch = true;
                    double d = Convert.ToDouble(cellVal);
                    if (fn == "MINIFS")
                    {
                        if (d < result) result = d;
                    }
                    else // MAXIFS
                    {
                        if (d > result) result = d;
                    }
                }
            }
        }

        // If no cell satisfies the criteria, Excel returns #VALUE!; we use NaN here
        data.CalculatedValue = anyMatch ? (object)result : double.NaN;
    }

    private bool IsNumeric(object val)
    {
        return val is double || val is int || val is long || val is float || val is decimal;
    }

    // Simple criteria evaluator supporting operators >, >=, <, <=, =, <> and plain equality
    private bool EvaluateCriteria(object cellValue, object criteria)
    {
        if (criteria == null)
            return false;

        if (criteria is string critStr)
        {
            critStr = critStr.Trim();

            if (critStr.StartsWith(">="))
                return Compare(cellValue, critStr.Substring(2)) >= 0;
            if (critStr.StartsWith("<="))
                return Compare(cellValue, critStr.Substring(2)) <= 0;
            if (critStr.StartsWith("<>"))
                return Compare(cellValue, critStr.Substring(2)) != 0;
            if (critStr.StartsWith(">"))
                return Compare(cellValue, critStr.Substring(1)) > 0;
            if (critStr.StartsWith("<"))
                return Compare(cellValue, critStr.Substring(1)) < 0;
            if (critStr.StartsWith("="))
                return Compare(cellValue, critStr.Substring(1)) == 0;

            // No operator -> equality
            return Compare(cellValue, critStr) == 0;
        }

        // Numeric criteria -> equality
        return Compare(cellValue, criteria) == 0;
    }

    private int Compare(object a, object b)
    {
        double da = Convert.ToDouble(a);
        double db = Convert.ToDouble(b);
        return da.CompareTo(db);
    }
}

class Program
{
    static void Main()
    {
        // Load the workbook that contains MINIFS / MAXIFS formulas
        string inputPath = "input.xlsx";
        Workbook wb = new Workbook(inputPath);

        // Attach the custom engine
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = new MinMaxIfsEngine()
        };

        // Calculate all formulas (our engine will handle MINIFS and MAXIFS)
        wb.CalculateFormula(opts);

        // Save the workbook with calculated results
        wb.Save("output.xlsx");
    }
}