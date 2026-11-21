// Title: Programmatically replace deprecated TEXTJOIN formulas with CONCATENATE in Excel files using Aspose.Cells for .NET
// Common Searches: C# code to convert TEXTJOIN formulas to CONCATENATE with Aspose.Cells | how to update Excel workbook formulas from TEXTJOIN to CONCATENATE using Aspose.Cells .NET | replace TEXTJOIN function in existing .xlsx files via Aspose.Cells API | regex based transformation of TEXTJOIN to CONCATENATE in C# | migrate legacy Excel formulas for older version compatibility using Aspose.Cells
// Tags: Aspose.Cells formula replacement | TEXTJOIN to CONCATENATE migration | C# regex for Excel formula conversion | handle deprecated Excel functions .NET | parse and split Excel arguments C#

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, scans every worksheet and cell for formulas containing the TEXTJOIN function, captures its delimiter, ignore‑empty flag, and arguments using a regular expression, splits the argument list while respecting nested parentheses, builds an equivalent CONCATENATE formula that inserts the delimiter between arguments (ignoring the ignore‑empty flag), replaces the original formula, and saves the updated workbook.
class TextJoinReplacer
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Regular expression to capture TEXTJOIN arguments:
        // Group 1: delimiter (e.g., \", \" or \",\")
        // Group 2: ignore_empty flag (TRUE/FALSE)
        // Group 3: the remaining arguments (could be a range or list of values)
        Regex textJoinRegex = new Regex(@"TEXTJOIN\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*(.+?)\)", RegexOptions.IgnoreCase);

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!cell.IsFormula) continue;

                string formula = cell.Formula;

                // Check if the formula contains TEXTJOIN
                if (!formula.Contains("TEXTJOIN", StringComparison.OrdinalIgnoreCase)) continue;

                // Replace each TEXTJOIN occurrence with an equivalent CONCATENATE formula
                string newFormula = textJoinRegex.Replace(formula, match =>
                {
                    string delimiter = match.Groups[1].Value.Trim();
                    string ignoreEmpty = match.Groups[2].Value.Trim();
                    string arguments = match.Groups[3].Value.Trim();

                    // Split arguments by commas that are not inside parentheses
                    var argsList = SplitArguments(arguments);

                    // Build CONCATENATE formula inserting the delimiter between arguments
                    // Note: This simple conversion does not handle ignore_empty = TRUE.
                    // For full compatibility, additional logic would be required.
                    string concatFormula = "CONCATENATE(";
                    for (int i = 0; i < argsList.Length; i++)
                    {
                        concatFormula += argsList[i];
                        if (i < argsList.Length - 1)
                        {
                            concatFormula += $", {delimiter}, ";
                        }
                    }
                    concatFormula += ")";

                    return concatFormula;
                });

                // Assign the transformed formula back to the cell
                cell.Formula = newFormula;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("Output.xlsx");
    }

    // Helper method to split function arguments while respecting nested parentheses
    private static string[] SplitArguments(string args)
    {
        var result = new System.Collections.Generic.List<string>();
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
        // Add the last argument
        result.Add(args.Substring(start).Trim());
        return result.ToArray();
    }
}
