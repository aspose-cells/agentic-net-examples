// Title: Replace legacy CONCATENATE formulas with the modern CONCAT operator in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Scan every worksheet in a .xlsx file with Aspose.Cells, locate cells whose formula contains CONCATENATE, and rewrite the formula to use CONCAT. | Write C# code that iterates through the used cell range, performs a case‑insensitive replacement of the CONCATENATE function with CONCAT, and saves the updated workbook.
// Common Searches: Aspose.Cells C# replace CONCATENATE function in all formulas | How to update Excel formulas from CONCATENATE to CONCAT programmatically | Bulk modify formulas in .xlsx using Aspose.Cells .NET | Iterate through worksheets and change legacy functions with Aspose.Cells
// Tags: replace CONCATENATE with CONCAT Aspose.Cells C# | bulk formula update Excel .xlsx Aspose.Cells | iterate worksheets modify formulas .NET | case-insensitive function rename Aspose.Cells | modernize Excel formulas Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook, iterates through each worksheet and its used cell range, detects formulas that contain the legacy CONCATENATE function, replaces the function name with the modern CONCAT operator in a case‑insensitive way, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        var workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula uses the old CONCATENATE function
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace CONCATENATE with the modern CONCAT function
                            string newFormula = formula.Replace("CONCATENATE", "CONCAT", StringComparison.OrdinalIgnoreCase);
                            cell.Formula = newFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
