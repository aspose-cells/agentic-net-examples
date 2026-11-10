// Title: Replace deprecated DATEVALUE function with DATE literals in Excel formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace DATEVALUE function in workbook formulas | Convert DATEVALUE to DATE literals in Excel using Aspose.Cells .NET | How to update Excel formulas from DATEVALUE to DATE programmatically | Regex based transformation of DATEVALUE formulas with Aspose.Cells | Remove deprecated DATEVALUE calls from .xlsx files using C# Aspose.Cells
// Tags: date literal substitution Aspose.Cells | excel formula refactoring C# | regex driven formula editing Aspose.Cells | handling deprecated Excel functions .NET | workbook formula migration Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, scans every worksheet and cell for formulas containing the DATEVALUE function, uses a regular expression to capture year, month, and day components, replaces DATEVALUE with the DATE(year, month, day) literal, updates the cell formula only when a change occurs, and saves the modified workbook.
class ReplaceDateValue
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Regular expression to match DATEVALUE with a date string argument
        // Supports formats like "2023-01-15", "2023/01/15", with or without quotes
        Regex dateValueRegex = new Regex(
            @"DATEVALUE\(\s*""?(\d{4})[-/](\d{1,2})[-/](\d{1,2})""?\s*\)",
            RegexOptions.IgnoreCase);

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // If the formula contains DATEVALUE, replace it with DATE(...)
                    if (formula.IndexOf("DATEVALUE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string newFormula = dateValueRegex.Replace(formula, "DATE($1,$2,$3)");
                        // Update the cell formula only if a change occurred
                        if (!newFormula.Equals(formula, StringComparison.Ordinal))
                        {
                            cell.Formula = newFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
