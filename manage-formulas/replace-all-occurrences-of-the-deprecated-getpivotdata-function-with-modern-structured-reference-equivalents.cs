// Title: Replace deprecated GETPIVOTDATA formulas with PivotTable structured references in Excel using Aspose.Cells for .NET (C#)
// Common Searches: C# code to convert GETPIVOTDATA functions to structured reference syntax with Aspose.Cells | Aspose.Cells replace GETPIVOTDATA formula in existing workbook | How to update Excel pivot formulas from GETPIVOTDATA to PivotTable[Field] using .NET | Migrate deprecated GETPIVOTDATA calls to modern structured references in C# | Automate Excel formula transformation from GETPIVOTDATA to structured reference with Aspose.Cells
// Tags: GETPIVOTDATA to structured reference conversion Aspose.Cells | C# Excel pivot formula replacement | Aspose.Cells formula migration .NET | update workbook formulas programmatically | structured reference pivot table Aspose

using Aspose.Cells;
using System;
using System.Text.RegularExpressions;

// The example loads an Excel workbook with Aspose.Cells, scans all cells for formulas containing the GETPIVOTDATA function, extracts the field name via regex, replaces each call with a PivotTable structured reference such as PivotTable1[FieldName], and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Regex to locate GETPIVOTDATA functions in formulas
        Regex getPivotRegex = new Regex(@"GETPIVOTDATA\(([^)]*)\)", RegexOptions.IgnoreCase);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula && cell.Formula.IndexOf("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string oldFormula = cell.Formula;

                    // Replace each GETPIVOTDATA call with a structured reference placeholder
                    // Example conversion:
                    // GETPIVOTDATA("Sum of Sales",$A$3) -> PivotTable1[Sum of Sales]
                    string newFormula = getPivotRegex.Replace(oldFormula, match =>
                    {
                        string args = match.Groups[1].Value;
                        // Split arguments; first argument is the data field name (quoted)
                        string[] parts = args.Split(new[] { ',' }, 2);
                        if (parts.Length > 0)
                        {
                            string fieldName = parts[0].Trim().Trim('\"');
                            // Use a generic pivot table name; adjust as needed for your workbook
                            return $"PivotTable1[{fieldName}]";
                        }
                        return string.Empty;
                    });

                    // Assign the transformed formula back to the cell
                    cell.Formula = newFormula;
                }
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save("output.xlsx");
    }
}
