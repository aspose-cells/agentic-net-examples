// Title: Replace every CONCATENATE function with CONCAT in all formulas of an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# program with Aspose.Cells that loads an .xlsx file, walks through each worksheet and cell, finds formulas containing the CONCATENATE function (ignoring case), swaps them to the CONCAT function, and saves the updated workbook. | Generate Aspose.Cells code that performs a case‑insensitive string replacement of "CONCATENATE" to "CONCAT" across all formula cells in a workbook and writes the result to a new file.
// Common Searches: Aspose.Cells C# replace CONCATENATE with CONCAT in all worksheet formulas | how to bulk update Excel formulas from CONCATENATE to CONCAT using .NET | programmatically change old Excel function names to new ones with Aspose.Cells | C# iterate through workbook cells and modify formula functions Aspose.Cells
// Tags: bulk replace Excel formula function Aspose.Cells | update CONCATENATE to CONCAT C# | iterate worksheets cells formula Aspose.Cells | case-insensitive formula string replacement .NET | save modified workbook as .xlsx Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// C# code that loads an Excel workbook with Aspose.Cells, iterates over every worksheet and cell, detects formulas containing the legacy CONCATENATE function (case‑insensitive), replaces it with the modern CONCAT function, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with actual input file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Check if the formula uses the old CONCATENATE function (case‑insensitive)
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace CONCATENATE with the modern CONCAT function
                            string updatedFormula = Regex.Replace(
                                formula,
                                "CONCATENATE",
                                "CONCAT",
                                RegexOptions.IgnoreCase);

                            // Assign the updated formula back to the cell
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with desired output file path)
        workbook.Save("output.xlsx");
    }
}
