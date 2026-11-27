// Title: Replace deprecated Excel functions with updated formulas using Aspose.Cells for .NET (C#)
// Common Searches: how to replace an old Excel function with a new one using Aspose.Cells C# | Aspose.Cells iterate all cells and update formulas containing a specific function | C# code to find and replace deprecated formula names in an Excel workbook | replace OLDFUNC with NEWFUNC in all worksheets using Aspose.Cells library | search and modify Excel formulas programmatically with Aspose.Cells .NET
// Tags: replace deprecated Excel function with Aspose.Cells | update formulas across all worksheets C# | regex based formula replacement .NET | iterate cells to modify formulas Aspose.Cells | detect and substitute function names in Excel formulas

using System;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, defines a deprecated function name and its recommended replacement, iterates through every worksheet and cell, identifies formulas that contain the old function (case‑insensitive), replaces the function using a regular‑expression replace, writes the updated formula back to each cell, and saves the modified workbook to a new file.
class ReplaceDeprecatedFormulas
{
    static void Main()
    {
        // Path to the input Excel file
        string inputPath = "input.xlsx";
        // Path to the output Excel file
        string outputPath = "output.xlsx";

        // Load the workbook (use the provided load rule)
        Workbook workbook = new Workbook(inputPath);

        // Define the deprecated function and its recommended alternative
        const string deprecatedFunction = "OLDFUNC";   // replace with actual deprecated function name
        const string alternativeFunction = "NEWFUNC"; // replace with the recommended alternative

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Check if the formula uses the deprecated function (case‑insensitive)
                    if (formula.IndexOf(deprecatedFunction, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the deprecated function with the alternative
                        string updatedFormula = ReplaceFunction(formula, deprecatedFunction, alternativeFunction);

                        // Assign the updated formula back to the cell
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Save the modified workbook (use the provided save rule)
        workbook.Save(outputPath);
    }

    private static string ReplaceFunction(string formula, string oldFunc, string newFunc)
    {
        // Build a case‑insensitive replacement while preserving the original case of the surrounding text
        // Use a simple replace with StringComparison since function names are distinct tokens
        return System.Text.RegularExpressions.Regex.Replace(
            formula,
            $@"\b{oldFunc}\b",
            newFunc,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }
}
