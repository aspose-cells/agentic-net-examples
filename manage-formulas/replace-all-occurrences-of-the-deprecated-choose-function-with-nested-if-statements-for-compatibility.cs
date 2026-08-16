// Title: Replace deprecated CHOOSE with nested IF using Aspose.Cells for .NET (C#)
// Description: C# sample that loads an Excel workbook with Aspose.Cells, scans every formula cell, detects the CHOOSE function, converts it to an equivalent nested IF expression via regex, updates the formula and saves the file. Ideal for bulk migration of legacy spreadsheets to Excel versions that no longer support CHOOSE.
// Keywords: Aspose.Cells | C# | CHOOSE function | nested IF | Excel formula conversion | replace CHOOSE | deprecated Excel function | bulk formula update | regex replace | Excel compatibility
// Common Searches: how to replace CHOOSE with IF in Excel using Aspose.Cells C# | convert CHOOSE formulas to nested IF programmatically | Aspose.Cells replace deprecated functions | C# code to change Excel formulas in bulk | regex CHOOSE to IF Aspose.Cells example
// Developer Intent: Transform all CHOOSE formulas in a workbook into nested IF statements for compatibility.
// Use Cases: Modernize legacy spreadsheets that rely on the CHOOSE function before distribution. | Automate large‑scale formula migration across multiple worksheets or workbooks. | Prepare Excel files for environments where CHOOSE is unsupported, such as older Office versions or third‑party parsers.
// AI Prompts: Generate a C# method with Aspose.Cells that replaces any CHOOSE call with a nested IF, handling any number of arguments. | Enhance the ReplaceChooseWithIf function to correctly parse arguments containing commas inside nested parentheses. | Write unit tests (e.g., using NUnit) that verify CHOOSE‑to‑IF conversion for different index expressions and argument counts.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsChooseReplacement
{
    // C# sample that loads an Excel workbook with Aspose.Cells, scans every formula cell, detects the CHOOSE function, converts it to an equivalent nested IF expression via regex, updates the formula and saves the file. Ideal for bulk migration of legacy spreadsheets to Excel versions that no longer support CHOOSE.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string formula = cell.Formula;

                    // Detect the deprecated CHOOSE function (case‑insensitive)
                    if (formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace CHOOSE with nested IF statements
                        string newFormula = ReplaceChooseWithIf(formula);
                        cell.Formula = newFormula;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }

        private static string ReplaceChooseWithIf(string formula)
        {
            // Regular expression to capture the argument list of CHOOSE
            // It assumes that arguments themselves do not contain commas inside parentheses.
            var choosePattern = new Regex(@"CHOOSE\s*\(([^)]*)\)", RegexOptions.IgnoreCase);
            var match = choosePattern.Match(formula);
            if (!match.Success)
                return formula; // No valid CHOOSE found; return original

            // Split the captured arguments by commas
            string[] args = match.Groups[1].Value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (args.Length < 2)
                return formula; // Not enough arguments to transform

            // Trim whitespace from each argument
            for (int i = 0; i < args.Length; i++)
                args[i] = args[i].Trim();

            string indexExpr = args[0]; // The index expression (e.g., A1 or 2)

            // Build nested IF expression
            string nestedIf = args[args.Length - 1]; // Default value (last argument)
            for (int i = args.Length - 2; i >= 1; i--)
            {
                // IF(index=position, value, previousNestedIf)
                nestedIf = $"IF({indexExpr}={i}, {args[i]}, {nestedIf})";
            }

            // Replace the original CHOOSE call with the new IF expression
            string newFormula = choosePattern.Replace(formula, nestedIf);
            // Ensure the formula starts with '=' (Aspose.Cells expects it)
            if (!newFormula.StartsWith("="))
                newFormula = "=" + newFormula;

            return newFormula;
        }
    }
}
