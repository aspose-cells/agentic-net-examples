using System;
using System.Collections.Generic;
using Aspose.Cells;

class ReplaceChooseWithIf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Sample worksheet with a CHOOSE formula
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(2); // index argument
        sheet.Cells["B1"].Formula = "=CHOOSE(A1, \"One\", \"Two\", \"Three\")";

        // Iterate through all worksheets and cells
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;
            foreach (Cell cell in cells)
            {
                string formula = cell.Formula;
                if (!string.IsNullOrEmpty(formula) &&
                    formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace CHOOSE with nested IF
                    string newFormula = ReplaceChoose(formula);
                    // Set the new formula (using default parse options)
                    cell.SetFormula(newFormula, new FormulaParseOptions());
                }
            }
        }

        // Save the updated workbook
        workbook.Save("ChooseReplaced.xlsx");
    }

    // Converts a CHOOSE function to an equivalent nested IF expression
    static string ReplaceChoose(string formula)
    {
        int start = formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase);
        if (start < 0) return formula;

        int openParen = formula.IndexOf('(', start);
        if (openParen < 0) return formula;

        int closeParen = FindMatchingParen(formula, openParen);
        if (closeParen < 0) return formula;

        // Extract the argument list inside CHOOSE(...)
        string argsContent = formula.Substring(openParen + 1, closeParen - openParen - 1);
        string[] parts = SplitArgs(argsContent);
        if (parts.Length < 2) return formula; // need at least index + one value

        string indexExpr = parts[0].Trim(); // the index expression

        // Build nested IF from the last value upwards
        string nestedIf = "";
        for (int i = parts.Length - 1; i >= 1; i--)
        {
            string value = parts[i].Trim();
            if (i == 1)
            {
                // Fallback when index is out of range – return the first value
                nestedIf = $"IF({indexExpr}=1,{value},{value})";
            }
            else
            {
                nestedIf = $"IF({indexExpr}={i},{value},{nestedIf})";
            }
        }

        // Replace the whole CHOOSE(...) segment with the generated IF expression
        string before = formula.Substring(0, start);
        string after = formula.Substring(closeParen + 1);
        return before + nestedIf + after;
    }

    // Finds the position of the matching closing parenthesis for '(' at startPos
    static int FindMatchingParen(string s, int startPos)
    {
        int depth = 0;
        for (int i = startPos; i < s.Length; i++)
        {
            if (s[i] == '(') depth++;
            else if (s[i] == ')')
            {
                depth--;
                if (depth == 0) return i;
            }
        }
        return -1; // not found
    }

    // Splits a comma‑separated argument string while respecting nested parentheses
    static string[] SplitArgs(string args)
    {
        List<string> list = new List<string>();
        int last = 0;
        int depth = 0;
        for (int i = 0; i < args.Length; i++)
        {
            char c = args[i];
            if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0)
            {
                list.Add(args.Substring(last, i - last));
                last = i + 1;
            }
        }
        list.Add(args.Substring(last));
        return list.ToArray();
    }
}