// Title: Convert deprecated CHOOSE formulas to nested IF with Aspose.Cells for .NET (C#)
// Description: A C# example that scans a workbook, detects every CHOOSE function, rewrites it as an equivalent nested IF expression using regular‑expression parsing, recalculates the sheet and saves the updated file—ensuring compatibility with older Excel versions.
// Keywords: Aspose.Cells CHOOSE replacement | C# convert CHOOSE to IF | nested IF formula generation | bulk Excel formula update | legacy Excel compatibility | .NET Excel automation | Excel formula migration | Aspose.Cells API example | US developers | global spreadsheet processing
// Common Searches: replace CHOOSE with IF Aspose.Cells C# | convert Excel CHOOSE function to nested IF programmatically | bulk formula conversion Aspose.Cells .NET | how to rewrite CHOOSE formulas in C# | Aspose.Cells example for updating formulas
// Developer Intent: Automatically rewrite all CHOOSE formulas in a workbook to nested IF statements for broader Excel compatibility.
// Use Cases: Modernize legacy spreadsheets that rely on the now‑deprecated CHOOSE function. | Perform bulk formula migration across multiple worksheets before distribution. | Validate conversion by recalculating the workbook and persisting the corrected version.
// AI Prompts: Create a C# method that parses any CHOOSE function (any number of arguments) and returns a nested IF formula, correctly handling commas inside quoted strings. | Write Aspose.Cells code to iterate through every cell in a workbook, replace CHOOSE formulas using the conversion method, recalculate all formulas, and save the workbook. | Explain how to adjust the regular expression to safely split CHOOSE arguments when they contain escaped quotes or embedded commas.

using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsChooseReplacement
{
    // A C# example that scans a workbook, detects every CHOOSE function, rewrites it as an equivalent nested IF expression using regular‑expression parsing, recalculates the sheet and saves the updated file—ensuring compatibility with older Excel versions.
    class Program
    {
        // Converts a CHOOSE function in a formula to nested IF statements.
        // Example: =CHOOSE(A1,"One","Two","Three")
        // becomes: =IF(A1=1,"One",IF(A1=2,"Two","Three"))
        static string ConvertChooseToIf(string formula)
        {
            // Find the CHOOSE function (case‑insensitive)
            var match = Regex.Match(formula, @"CHOOSE\s*\(([^)]*)\)", RegexOptions.IgnoreCase);
            if (!match.Success) return formula; // No CHOOSE found

            // Extract the inner arguments
            string argsContent = match.Groups[1].Value;
            // Split by commas, but keep commas inside quotes intact
            var args = Regex.Split(argsContent, @",(?![^""]*""\s*,)").Select(s => s.Trim()).ToArray();

            if (args.Length < 2) return formula; // Not enough arguments

            string indexExpr = args[0]; // The index expression (could be a cell reference or number)
            // Build nested IFs from the remaining arguments
            StringBuilder sb = new StringBuilder();
            sb.Append("IF(");
            sb.Append(indexExpr);
            sb.Append("=1,");
            sb.Append(args[1]);

            for (int i = 2; i < args.Length; i++)
            {
                sb.Append(",IF(");
                sb.Append(indexExpr);
                sb.Append("=");
                sb.Append(i);
                sb.Append(",");
                sb.Append(args[i]);
            }

            // Close the opened parentheses
            sb.Append(new string(')', args.Length - 1));

            // Replace the original CHOOSE call with the generated IF chain
            string newFormula = Regex.Replace(formula,
                @"CHOOSE\s*\([^\)]*\)",
                sb.ToString(),
                RegexOptions.IgnoreCase);

            return newFormula;
        }

        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data for the index argument
            cells["A1"].PutValue(2); // Index = 2

            // Cell with a CHOOSE formula (to be replaced)
            cells["B1"].Formula = "=CHOOSE(A1,\"One\",\"Two\",\"Three\")";

            // Iterate through all cells and replace CHOOSE with nested IF
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Formula.IndexOf("CHOOSE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string original = cell.Formula;
                    string replaced = ConvertChooseToIf(original);
                    cell.Formula = replaced;
                }
            }

            // Calculate formulas to verify the replacement works (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChooseReplaced.xlsx");
        }
    }
}
