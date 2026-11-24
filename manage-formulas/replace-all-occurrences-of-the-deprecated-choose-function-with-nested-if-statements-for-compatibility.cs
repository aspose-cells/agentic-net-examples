// Title: Replace deprecated CHOOSE function with nested IF formulas in Excel files using Aspose.Cells for .NET
// Common Searches: how to convert CHOOSE formulas to IF in Aspose.Cells C# | Aspose.Cells replace CHOOSE with nested IF in a workbook | C# program to update Excel formulas from CHOOSE to IF | migrate deprecated CHOOSE function to IF using Aspose.Cells API | batch convert CHOOSE formulas to IF statements in .xlsx with Aspose.Cells | automatically change Excel CHOOSE to IF in C# code
// Tags: CHOOSE to IF formula conversion Aspose.Cells | nested IF generation for Excel formulas C# | deprecated Excel function update programmatically | cell formula transformation Aspose.Cells .NET | batch workbook formula processing C#

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through all worksheets and cells, detects formulas that use the CHOOSE function, converts each occurrence to an equivalent nested IF expression using regular‑expression parsing, replaces the original formula, and saves the updated workbook to a new file.
class ChooseToIfConverter
{
    // Convert a CHOOSE formula to an equivalent nested IF formula.
    // Example: =CHOOSE(A1, "One", "Two", "Three")
    // becomes: =IF(A1=1,"One",IF(A1=2,"Two","Three"))
    static string ConvertChooseToIf(string formula)
    {
        // Remove leading '=' if present
        string expr = formula.StartsWith("=") ? formula.Substring(1) : formula;

        // Find CHOOSE function and its arguments
        var match = Regex.Match(expr, @"CHOOSE\s*\(([^)]*)\)", RegexOptions.IgnoreCase);
        if (!match.Success)
            return formula; // no CHOOSE found

        // Split arguments by commas (simple split, assumes no commas inside quoted strings)
        string[] args = match.Groups[1].Value.Split(',');

        if (args.Length < 2)
            return formula; // not enough arguments

        string indexExpr = args[0].Trim(); // the index expression (e.g., A1 or 2)

        // Build nested IF expression
        StringBuilder ifBuilder = new StringBuilder();
        for (int i = 1; i < args.Length; i++)
        {
            string value = args[i].Trim();

            if (i == args.Length - 1)
            {
                // Last value: no further nesting
                ifBuilder.Append($"IF({indexExpr}={i}, {value}, {value})");
            }
            else
            {
                ifBuilder.Append($"IF({indexExpr}={i}, {value}, ");
            }
        }

        // Close opened parentheses
        for (int i = 1; i < args.Length - 1; i++)
        {
            ifBuilder.Append(")");
        }

        // Replace original CHOOSE call with the new IF expression
        string newExpr = Regex.Replace(expr, @"CHOOSE\s*\([^)]+\)", ifBuilder.ToString(), RegexOptions.IgnoreCase);
        return "=" + newExpr;
    }

    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if cell contains a formula with CHOOSE
                    if (cell.IsFormula && cell.Formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string originalFormula = cell.Formula;
                        string updatedFormula = ConvertChooseToIf(originalFormula);
                        if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                        {
                            cell.Formula = updatedFormula; // assign new IF‑based formula
                        }
                    }
                }
            }

            // Save modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
