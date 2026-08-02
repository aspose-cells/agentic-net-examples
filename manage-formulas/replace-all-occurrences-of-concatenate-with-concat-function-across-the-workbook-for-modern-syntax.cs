// Title: Replace CONCATENATE with CONCAT in Excel formulas using Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook with Aspose.Cells, scans every worksheet, finds formulas containing the legacy CONCATENATE function, swaps them for the modern CONCAT syntax, recalculates all formulas, and saves the updated file.
// Keywords: Aspose.Cells replace CONCATENATE | C# update Excel formula CONCAT | convert CONCATENATE to CONCAT programmatically | Aspose.Cells formula edit example | Excel modern functions .NET | batch replace Excel functions
// Common Searches: Aspose.Cells replace CONCATENATE with CONCAT | C# code to change Excel formulas from CONCATENATE to CONCAT | how to update legacy Excel functions using Aspose.Cells | recalculate formulas after function replacement .NET | bulk edit Excel formulas programmatically
// Developer Intent: Swap every CONCATENATE call for CONCAT in all formula cells of a workbook.
// Use Cases: Modernize legacy workbooks for compatibility with Excel 2016+ | Automate bulk formula upgrades in generated reports | Prepare files for cloud‑based analytics platforms that require current functions | Ensure consistent formula syntax across multinational teams (US, EU, APAC)
// AI Prompts: Generate C# code with Aspose.Cells that finds and replaces CONCATENATE with CONCAT in all worksheets, then recalculates and saves the workbook. | Show an optimized approach using Worksheet.Cells.Find to target only cells containing CONCATENATE before applying the replacement. | Explain how to preserve the original behavior when CONCATENATE has a variable number of arguments after converting to CONCAT.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// C# sample that loads an Excel workbook with Aspose.Cells, scans every worksheet, finds formulas containing the legacy CONCATENATE function, swaps them for the modern CONCAT syntax, recalculates all formulas, and saves the updated file.
class ReplaceConcatenate
{
    static void Main()
    {
        // Load the workbook that contains formulas using CONCATENATE
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each cell that has data in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Check if the formula uses the old CONCATENATE function (case‑insensitive)
                    if (formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace all occurrences of CONCATENATE with the modern CONCAT function
                        string newFormula = Regex.Replace(
                            formula,
                            @"CONCATENATE",
                            "CONCAT",
                            RegexOptions.IgnoreCase);

                        // Assign the updated formula back to the cell
                        cell.Formula = newFormula;
                    }
                }
            }
        }

        // Recalculate all formulas after the replacement
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
