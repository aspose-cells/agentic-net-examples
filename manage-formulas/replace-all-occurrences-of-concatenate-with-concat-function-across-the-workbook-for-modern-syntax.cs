// Title: Replace CONCATENATE with CONCAT in every Excel formula using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace CONCATENATE function in all worksheets | programmatically convert legacy CONCATENATE formulas to CONCAT in an Excel workbook | bulk update Excel formulas to modern CONCAT using Aspose.Cells .NET | C# code to change Excel formula functions from CONCATENATE to CONCAT across workbook | how to modernize Excel formulas with Aspose.Cells by swapping CONCATENATE for CONCAT
// Tags: CONCATENATE to CONCAT conversion Aspose.Cells | bulk formula replacement .NET | Excel formula modernization C# | regex based formula edit Aspose.Cells | iterate all cells Aspose.Cells

using Aspose.Cells;
using System;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet and cell, detects formulas that use the legacy CONCATENATE function (case‑insensitive), replaces the function name with the modern CONCAT using a regular expression, writes the updated formula back to each cell, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Detect usage of the old CONCATENATE function (case‑insensitive)
                    if (formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the function name with the modern CONCAT
                        string newFormula = Regex.Replace(
                            formula,
                            @"\bCONCATENATE\b",
                            "CONCAT",
                            RegexOptions.IgnoreCase);

                        // Update the cell with the new formula
                        cell.Formula = newFormula;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
