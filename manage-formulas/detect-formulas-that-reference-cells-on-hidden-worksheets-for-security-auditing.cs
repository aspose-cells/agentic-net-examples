// Title: Detect formulas that reference hidden worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, identifies hidden worksheets, and returns a list of cells whose formulas reference those hidden sheets. | Refactor the sample to output the audit results as JSON for easy consumption by a security reporting pipeline. | Extend the solution to ignore formulas that reference hidden sheets only through named ranges while still reporting direct sheet references.
// Common Searches: how to find Excel formulas that point to hidden sheets using Aspose.Cells C# | C# audit workbook for hidden worksheet references in formulas | list cells with formulas referencing hidden worksheets Aspose.Cells .NET | security scan Excel file for hidden sheet links using Aspose.Cells | detect hidden sheet formula references programmatically in .NET
// Tags: scan workbook for hidden sheet formula references | Aspose.Cells detect hidden worksheet links | C# audit Excel formulas referencing hidden sheets | regex extract sheet name from formula Aspose.Cells | security audit hidden worksheet references .NET

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads a workbook with Aspose.Cells, gathers the names of hidden worksheets, iterates all formula cells, uses a regular expression to locate sheet references, and records any cell whose formula points to a hidden sheet, then outputs the findings.
class Program
{
    static void Main(string[] args)
    {
        // Expect the first argument to be the path of the workbook to audit.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <workbookPath>");
            return;
        }

        string workbookPath = args[0];

        // Load the workbook.
        Workbook workbook = new Workbook(workbookPath);

        // Collect names of hidden worksheets.
        HashSet<string> hiddenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // In Aspose.Cells, Worksheet.IsVisible indicates whether the sheet is visible.
            if (!ws.IsVisible)
            {
                hiddenSheetNames.Add(ws.Name);
            }
        }

        // Prepare a regex to capture sheet names in formulas.
        // It matches patterns like Sheet1!A1 or 'My Hidden Sheet'!B2
        Regex sheetRefRegex = new Regex(@"(?i)(?:'(?<sheet>[^']+)'|(?<sheet>[^'!\s]+))!", RegexOptions.Compiled);

        // List to store findings.
        List<string> findings = new List<string>();

        // Scan all cells with formulas.
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Cells cells = ws.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Find all sheet references in the formula.
                    foreach (Match match in sheetRefRegex.Matches(formula))
                    {
                        string referencedSheet = match.Groups["sheet"].Value;
                        if (hiddenSheetNames.Contains(referencedSheet))
                        {
                            // Record the occurrence.
                            findings.Add($"Cell {ws.Name}!{cell.Name} contains a formula referencing hidden sheet '{referencedSheet}'. Formula: {formula}");
                            // No need to check further references for this cell.
                            break;
                        }
                    }
                }
            }
        }

        // Output the audit results.
        if (findings.Count == 0)
        {
            Console.WriteLine("No formulas referencing hidden worksheets were found.");
        }
        else
        {
            Console.WriteLine("Formulas referencing hidden worksheets:");
            foreach (string line in findings)
            {
                Console.WriteLine(line);
            }
        }
    }
}
