// Title: Programmatically replace deprecated SUMIF formulas with SUMIFS across all worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace SUMIF with SUMIFS in entire workbook | convert SUMIF formulas to SUMIFS using .NET API | bulk update Excel formulas from SUMIF to SUMIFS with Aspose.Cells | iterate through worksheets and modify formulas in C# | parse and transform SUMIF to SUMIFS programmatically
// Tags: SUMIF to SUMIFS conversion Aspose.Cells | bulk Excel formula update .NET | worksheet iteration replace formulas C# | deprecated function handling Aspose.Cells | custom formula parsing SUMIFS transformation

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, walks through every worksheet and cell, detects any formula containing the SUMIF function, converts it to an equivalent SUMIFS formula using custom parsing logic, updates the cell's formula, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;
                    if (!string.IsNullOrEmpty(formula) &&
                        formula.IndexOf("SUMIF", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace deprecated SUMIF with SUMIFS
                        string newFormula = ReplaceSumIfWithSumifs(formula);
                        cell.Formula = newFormula;
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }

    // Converts a SUMIF formula to an equivalent SUMIFS formula
    static string ReplaceSumIfWithSumifs(string formula)
    {
        // Locate the SUMIF function name
        int funcPos = formula.IndexOf("SUMIF", StringComparison.OrdinalIgnoreCase);
        if (funcPos < 0) return formula;

        // Find the opening parenthesis of the function
        int openParen = formula.IndexOf('(', funcPos);
        if (openParen < 0) return formula;

        // Find the matching closing parenthesis
        int closeParen = FindMatchingParen(formula, openParen);
        if (closeParen < 0) return formula;

        // Extract the argument string inside the parentheses
        string argsContent = formula.Substring(openParen + 1, closeParen - openParen - 1);

        // Split arguments while respecting nested parentheses
        List<string> args = SplitArguments(argsContent);

        // SUMIF can have 2 or 3 arguments:
        //   SUMIF(range, criteria)               -> SUMIFS(range, range, criteria)
        //   SUMIF(range, criteria, sum_range)    -> SUMIFS(sum_range, range, criteria)
        string range, criteria, sumRange;

        if (args.Count == 2)
        {
            range = args[0];
            criteria = args[1];
            sumRange = range; // When sum_range is omitted, it defaults to range
        }
        else if (args.Count >= 3)
        {
            range = args[0];
            criteria = args[1];
            sumRange = args[2];
        }
        else
        {
            // Unexpected argument count; return original formula
            return formula;
        }

        // Build the new SUMIFS argument list: sum_range, range, criteria
        string newArgs = $"{sumRange},{range},{criteria}";

        // Construct the new formula string
        string newFormula = formula.Substring(0, funcPos) + "SUMIFS(" + newArgs + ")" + formula.Substring(closeParen + 1);
        return newFormula;
    }

    // Finds the position of the matching closing parenthesis for a given opening parenthesis
    static int FindMatchingParen(string text, int openPos)
    {
        int depth = 0;
        for (int i = openPos; i < text.Length; i++)
        {
            if (text[i] == '(') depth++;
            else if (text[i] == ')')
            {
                depth--;
                if (depth == 0) return i;
            }
        }
        return -1; // No matching parenthesis found
    }

    // Splits a comma‑separated argument string, ignoring commas inside nested parentheses
    static List<string> SplitArguments(string args)
    {
        var result = new List<string>();
        int start = 0;
        int depth = 0;

        for (int i = 0; i < args.Length; i++)
        {
            char c = args[i];
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0)
            {
                result.Add(args.Substring(start, i - start).Trim());
                start = i + 1;
            }
        }

        // Add the final argument
        if (start < args.Length)
            result.Add(args.Substring(start).Trim());

        return result;
    }
}
