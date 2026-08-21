// Title: C# – Replace POWER(x,y) with ^ operator in all Excel formulas using Aspose.Cells
// Description: Loads an Excel workbook, scans every worksheet and cell, uses a regular expression to change each POWER(x,y) formula to x^y, updates the cell, recalculates all formulas, and saves the modified file.
// Keywords: Aspose.Cells | C# formula replace | POWER to ^ operator | regex Excel formula | bulk formula update | recalculate workbook | Excel exponentiation | replace POWER function | .NET Excel automation | Aspose.Cells formula editing
// Common Searches: Aspose.Cells replace POWER function with caret | C# change POWER(x,y) to ^ in Excel workbook | bulk formula replacement using Aspose.Cells | regex replace POWER in Excel sheets .NET | recalculate workbook after formula changes Aspose.Cells | convert Excel POWER to ^ operator programmatically
// Developer Intent: Automatically convert every POWER(x,y) occurrence to the ^ exponentiation syntax across all worksheets in an Excel workbook using Aspose.Cells.
// Use Cases: Standardize formulas before sharing a workbook by switching from POWER to the native ^ operator. | Prepare files for platforms that lack the POWER function, ensuring calculations remain correct after conversion. | Automate large‑scale formula updates in a server‑side process while automatically recalculating results.
// AI Prompts: Write C# code with Aspose.Cells that finds and replaces POWER(x,y) with x^y in all formulas, then recalculates the workbook. | Provide a regex pattern and loop structure for Aspose.Cells to convert POWER functions to the ^ operator without altering cell formatting. | Explain how to handle nested POWER calls and maintain correct precedence when swapping them for the ^ exponentiation operator in Aspose.Cells.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPowerReplacement
{
    // Loads an Excel workbook, scans every worksheet and cell, uses a regular expression to change each POWER(x,y) formula to x^y, updates the cell, recalculates all formulas, and saves the modified file.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Replace POWER(x,y) with x^y using a regular expression
                        string originalFormula = cell.Formula;
                        string updatedFormula = Regex.Replace(
                            originalFormula,
                            @"POWER\(\s*([^,]+)\s*,\s*([^\)]+)\s*\)",
                            "$1^$2",
                            RegexOptions.IgnoreCase);

                        // If a replacement occurred, set the new formula
                        if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                        {
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after modification
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
