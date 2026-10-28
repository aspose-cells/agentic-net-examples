// Title: Build a cached formula evaluation engine in C# using Aspose.Cells to efficiently compute and replace Excel formulas
// Common Searches: C# Aspose.Cells cache identical sub‑expressions when evaluating Excel formulas | how to create a custom formula evaluator that avoids redundant calculations in Aspose.Cells | replace Excel cell formulas with computed values using a cached engine in C# | recursive evaluation of cell references with caching in Aspose.Cells | optimize performance of bulk formula calculation in Aspose.Cells C#
// Tags: cached formula evaluation Aspose.Cells | recursive cell reference resolution C# | DataTable.Compute expression parsing | replace Excel formulas with static values C# | optimize Excel calculation performance Aspose.Cells | avoid redundant subexpression computation C#

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example defines a CachedFormulaEngine that evaluates Excel formulas on a Worksheet, substitutes cell references (including nested formulas) with their numeric values, computes the arithmetic using DataTable.Compute, and caches results of identical sub‑expressions to prevent repeated work. The program loads a workbook, runs the engine on each formula cell, replaces the formula with the calculated value, and saves the updated file.
class CachedFormulaEngine
{
    private readonly Worksheet _worksheet;
    private readonly Dictionary<string, double> _cache = new Dictionary<string, double>();

    public CachedFormulaEngine(Worksheet worksheet)
    {
        _worksheet = worksheet;
    }

    // Public entry point for evaluating a formula string
    public double Evaluate(string formula)
    {
        // Remove leading '=' if present (Aspose.Cells stores formulas with '=')
        if (formula.StartsWith("="))
            formula = formula.Substring(1);

        return EvaluateExpression(formula);
    }

    // Core evaluator that caches results of identical sub‑expressions
    private double EvaluateExpression(string expr)
    {
        expr = expr.Trim();

        // Return cached result if this exact expression was evaluated before
        if (_cache.TryGetValue(expr, out double cachedResult))
            return cachedResult;

        // Replace all cell references (e.g., A1, B12) with their numeric values
        string replacedExpr = Regex.Replace(expr, @"[A-Za-z]+\d+", match =>
        {
            string address = match.Value;
            Cell refCell = _worksheet.Cells[address];

            // If the referenced cell itself contains a formula, evaluate it recursively
            if (refCell.IsFormula)
                return Evaluate(refCell.Formula).ToString();

            // Otherwise use the cell's current value (default to 0 if empty)
            return refCell.Value?.ToString() ?? "0";
        });

        // Use DataTable.Compute to evaluate the arithmetic expression
        double result;
        try
        {
            var dt = new DataTable();
            result = Convert.ToDouble(dt.Compute(replacedExpr, string.Empty));
        }
        catch
        {
            // In case of an invalid expression, return NaN
            result = double.NaN;
        }

        // Cache the computed result for future identical sub‑expressions
        _cache[expr] = result;
        return result;
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load an existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Instantiate the custom cached formula engine
        var engine = new CachedFormulaEngine(sheet);

        // Iterate over all cells in the worksheet
        foreach (Cell cell in sheet.Cells)
        {
            // Process only cells that contain formulas
            if (cell.IsFormula)
            {
                try
                {
                    // Evaluate the formula using the caching engine
                    double value = engine.Evaluate(cell.Formula);

                    // Replace the formula with the calculated value (PutValue clears the formula)
                    cell.PutValue(value);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error evaluating cell {cell.Name}: {ex.Message}");
                }
            }
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
