// Title: How to replace a deprecated Excel function with a new one in every formula across a workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, scans each worksheet's used range, finds formulas containing a specified deprecated function, replaces it with a new function name (case‑insensitive), and saves the updated file. | Generate a reusable method in C# that accepts an input file path, an old Excel function name, and a replacement function name, then updates all formula cells in the workbook using Aspose.Cells. | Create a utility class for Aspose.Cells that iterates through all cells in a workbook, performs a case‑insensitive string replace of an outdated function within formulas, and writes the changes back to a new file.
// Common Searches: Aspose.Cells replace old Excel function name in formulas C# | C# update deprecated Excel functions across all worksheets using Aspose.Cells | bulk modify formula functions in an Excel file with Aspose.Cells .NET | search and replace function names in Excel formulas programmatically Aspose.Cells | iterate used range and change Excel function name in workbook using Aspose.Cells
// Tags: replace deprecated function Aspose.Cells | bulk formula update .NET | scan used range Aspose.Cells | case-insensitive formula replace C# | excel workbook function substitution Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, defines the name of a deprecated function and its replacement, then iterates through every worksheet's used cell range. For each cell containing a formula, it performs a case‑insensitive replacement of the old function with the new one and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the deprecated function name and its recommended replacement
        const string deprecatedFunction = "OLD_FUNC";
        const string replacementFunction = "NEW_FUNC";

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of cells for the current sheet
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // If the formula uses the deprecated function, replace it
                        if (cell.Formula.IndexOf(deprecatedFunction, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string updatedFormula = cell.Formula.Replace(
                                deprecatedFunction,
                                replacementFunction,
                                StringComparison.OrdinalIgnoreCase);

                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
