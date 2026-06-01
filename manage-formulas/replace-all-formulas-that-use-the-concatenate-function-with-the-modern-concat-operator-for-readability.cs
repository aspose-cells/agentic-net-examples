using System;
using Aspose.Cells;

class ReplaceConcat
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet (adjust as needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Iterate through all cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Process only cells that contain a formula
            if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.StartsWith("="))
            {
                string formula = cell.Formula;

                // Detect the CONCATENATE function (case‑insensitive)
                if (formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Convert CONCATENATE(...) to the modern CONCAT operator (&)
                    string newFormula = ConvertConcatenateToConcatOperator(formula);

                    // Update the cell with the new formula
                    cell.Formula = newFormula;
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Helper method that rewrites a CONCATENATE formula using the '&' operator
    static string ConvertConcatenateToConcatOperator(string formula)
    {
        // Locate the CONCATENATE function within the formula
        int startIdx = formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase);
        int openParenIdx = formula.IndexOf('(', startIdx);
        int closeParenIdx = formula.LastIndexOf(')');

        // If the parentheses are not found correctly, return the original formula
        if (openParenIdx < 0 || closeParenIdx < 0 || closeParenIdx <= openParenIdx)
            return formula;

        // Extract the argument list inside CONCATENATE(...)
        string argsInside = formula.Substring(openParenIdx + 1, closeParenIdx - openParenIdx - 1);

        // Simple split on commas (works for most cases without nested commas)
        string[] args = argsInside.Split(',');

        // Trim whitespace from each argument
        for (int i = 0; i < args.Length; i++)
        {
            args[i] = args[i].Trim();
        }

        // Re‑join arguments using the '&' operator
        string concatenated = string.Join(" & ", args);

        // Reconstruct the formula, preserving any surrounding expression
        string before = formula.Substring(0, startIdx);
        string after = formula.Substring(closeParenIdx + 1);
        string result = before + concatenated + after;

        // Ensure the formula still starts with '='
        if (!result.StartsWith("="))
            result = "=" + result;

        return result;
    }
}