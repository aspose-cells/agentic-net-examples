// Title: How to replace the deprecated REPT function with REPEAT in every formula of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates all worksheets and cells, and substitutes the REPT function name with REPEAT in any formula. | Create a .NET method that scans a workbook's used range and updates formulas containing REPT to use the REPEAT function, preserving case. | Write a C# routine that uses a regular expression to replace the REPT keyword with REPEAT across all formula cells and then saves the workbook.
// Common Searches: Aspose.Cells C# replace REPT with REPEAT in all worksheet formulas | bulk update deprecated Excel functions using Aspose.Cells .NET | programmatically change function names in Excel formulas with Aspose.Cells | C# iterate through workbook cells to modify formula functions | replace Excel REPT function in existing file using Aspose.Cells library
// Tags: update REPT to REPEAT Aspose.Cells C# | bulk formula modification Excel .NET | regex formula substitution Aspose.Cells | iterate workbook cells Aspose.Cells | deprecated Excel function replacement programmatically

using Aspose.Cells;
using System;
using System.Text.RegularExpressions;

// The program loads an Excel workbook with Aspose.Cells, walks through every worksheet and cell in the used range, detects formulas that contain the deprecated REPT function, replaces the function name with REPEAT using a case‑insensitive regular expression, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection for the current worksheet
            Cells cells = sheet.Cells;

            // Determine the used range boundaries
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Loop through each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // Identify formulas that use the deprecated REPT function (case‑insensitive)
                        if (formula.IndexOf("REPT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Replace the REPT function name with REPEAT while preserving the rest of the formula
                            string updatedFormula = Regex.Replace(
                                formula,
                                @"\bREPT\b",
                                "REPEAT",
                                RegexOptions.IgnoreCase);

                            // Assign the corrected formula back to the cell
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }
        }

        // Save the modified workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
