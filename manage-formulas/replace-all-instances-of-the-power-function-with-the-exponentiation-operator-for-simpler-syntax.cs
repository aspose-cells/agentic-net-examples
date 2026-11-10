// Title: Replace Excel POWER() function with ^ exponent operator in formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change POWER function to ^ operator in an Excel file with Aspose.Cells | How to programmatically convert POWER(arg1,arg2) to (arg1)^(arg2) using Aspose.Cells .NET | Aspose.Cells replace Power formula with caret operator across all worksheets | Iterate cells and update Excel formulas from POWER to ^ in C# | Modify workbook formulas to use exponentiation operator with Aspose.Cells API
// Tags: Aspose.Cells exponent operator substitution | C# Excel formula transformation | bulk formula update in .xlsx via Aspose.Cells | custom formula parser for POWER replacement | Aspose.Cells workbook formula editing

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The example loads an Excel workbook, walks through every worksheet and cell, detects formulas that contain the POWER function, and replaces each occurrence with the equivalent caret (^) exponentiation syntax using custom parsing logic, then saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Loop through all cells that contain a formula
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Convert POWER functions to the ^ operator
                    string updatedFormula = ReplacePowerFunction(cell.Formula);
                    if (updatedFormula != cell.Formula)
                    {
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Replaces all occurrences of POWER(arg1,arg2) with (arg1)^(arg2)
    static string ReplacePowerFunction(string formula)
    {
        int searchPos = 0;
        while ((searchPos = formula.IndexOf("POWER(", searchPos, StringComparison.OrdinalIgnoreCase)) != -1)
        {
            int openParen = formula.IndexOf('(', searchPos);
            if (openParen == -1) break;

            int closeParen = FindMatchingParenthesis(formula, openParen);
            if (closeParen == -1) break;

            // Extract the argument list inside POWER(...)
            string args = formula.Substring(openParen + 1, closeParen - openParen - 1);
            string[] parts = SplitArguments(args);

            // Only replace when exactly two arguments are present
            if (parts.Length == 2)
            {
                string replacement = $"({parts[0]})^({parts[1]})";
                formula = formula.Substring(0, searchPos) + replacement + formula.Substring(closeParen + 1);
                // Move past the newly inserted text
                searchPos += replacement.Length;
            }
            else
            {
                // If parsing fails, skip this occurrence
                searchPos = closeParen + 1;
            }
        }
        return formula;
    }

    // Finds the index of the closing parenthesis that matches the opening at startPos
    static int FindMatchingParenthesis(string text, int startPos)
    {
        int depth = 0;
        for (int i = startPos; i < text.Length; i++)
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
    static string[] SplitArguments(string args)
    {
        List<string> result = new List<string>();
        int lastPos = 0;
        int depth = 0;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == '(') depth++;
            else if (args[i] == ')') depth--;
            else if (args[i] == ',' && depth == 0)
            {
                result.Add(args.Substring(lastPos, i - lastPos).Trim());
                lastPos = i + 1;
            }
        }
        result.Add(args.Substring(lastPos).Trim());
        return result.ToArray();
    }
}
