// Title: Convert deprecated SUMIF formulas to SUMIFS throughout an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells, scans every worksheet, detects cells containing the SUMIF function, and rewrites each formula to the equivalent SUMIFS syntax. | Write a .NET routine that applies a regular expression to replace SUMIF(range, criteria) and SUMIF(range, criteria, sum_range) with the proper SUMIFS form in all formula cells of a workbook. | Create a script that iterates over all cells in an Aspose.Cells workbook, transforms deprecated SUMIF usage into SUMIFS, and saves the updated workbook.
// Common Searches: Aspose.Cells replace SUMIF with SUMIFS in C# | how to update deprecated SUMIF formulas in an Excel file using .NET | bulk convert SUMIF to SUMIFS across all worksheets programmatically | regex formula replacement SUMIF to SUMIFS Aspose.Cells example | C# code to iterate through workbook cells and modify formulas
// Tags: replace SUMIF with SUMIFS Aspose.Cells | bulk formula transformation .NET | regex formula update Excel C# | iterate worksheets modify formulas Aspose | deprecated Excel function migration Aspose.Cells | update Excel formulas programmatically C#

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, iterates through every worksheet and formula cell, uses a regular expression to locate SUMIF calls, rewrites them as equivalent SUMIFS expressions, assigns the new formulas, and saves the modified file.
class ReplaceSumIfWithSumifs
{
    static void Main()
    {
        // Load the workbook (use the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Regular expression to find SUMIF functions and capture their arguments
        Regex sumIfRegex = new Regex(@"SUMIF\(([^()]*)\)", RegexOptions.IgnoreCase);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!cell.IsFormula) continue;

                string originalFormula = cell.Formula;

                // Replace each SUMIF occurrence with the equivalent SUMIFS
                string updatedFormula = sumIfRegex.Replace(originalFormula, match =>
                {
                    // Split the captured arguments by commas
                    string argsPart = match.Groups[1].Value;
                    string[] args = argsPart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    // Trim whitespace from each argument
                    for (int i = 0; i < args.Length; i++)
                        args[i] = args[i].Trim();

                    // SUMIF can have 2 or 3 arguments
                    // SUMIF(range, criteria)               -> SUMIFS(range, range, criteria)
                    // SUMIF(range, criteria, sum_range)    -> SUMIFS(sum_range, range, criteria)
                    string sumifsFormula;
                    if (args.Length == 2)
                    {
                        // range, criteria
                        string range = args[0];
                        string criteria = args[1];
                        sumifsFormula = $"SUMIFS({range}, {range}, {criteria})";
                    }
                    else if (args.Length == 3)
                    {
                        // range, criteria, sum_range
                        string range = args[0];
                        string criteria = args[1];
                        string sumRange = args[2];
                        sumifsFormula = $"SUMIFS({sumRange}, {range}, {criteria})";
                    }
                    else
                    {
                        // Unexpected number of arguments; keep original
                        return match.Value;
                    }

                    return sumifsFormula;
                });

                // If the formula changed, assign the new formula back to the cell
                if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                {
                    cell.Formula = updatedFormula;
                }
            }
        }

        // Save the modified workbook (use the provided save rule)
        workbook.Save("output.xlsx");
    }
}
