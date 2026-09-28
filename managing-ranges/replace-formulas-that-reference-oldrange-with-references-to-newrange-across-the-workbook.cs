// Title: Replace all occurrences of the named range 'OldRange' with 'NewRange' in formulas across an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet in a workbook and substitutes the named range 'OldRange' with 'NewRange' in any formula, then saves the file. | Generate a method that updates formulas containing a specific named range to a new name across all sheets using Aspose.Cells. | Create a script that loads an .xlsx file, replaces 'OldRange' references in formulas with 'NewRange', and writes the result to a new file.
// Common Searches: how to rename a named range in all formulas using Aspose.Cells C# | replace specific range name in Excel formulas programmatically .NET | bulk update named range references across worksheets Aspose.Cells | C# Aspose.Cells change formula text OldRange to NewRange
// Tags: replace named range in formulas Aspose.Cells | bulk formula text substitution .NET | iterate worksheets Aspose.Cells | named range rename Excel C# | update formula references Aspose.Cells

using Aspose.Cells;
using System;

// The program loads an Excel workbook, iterates through every worksheet and cell, replaces any formula that references the named range 'OldRange' with 'NewRange', and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection for the current worksheet
            Cells cells = sheet.Cells;

            // Loop through each cell in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    string formula = cell.Formula;

                    // Check if the formula references the old range name
                    if (!string.IsNullOrEmpty(formula) && formula.Contains("OldRange"))
                    {
                        // Replace "OldRange" with "NewRange" in the formula
                        string updatedFormula = formula.Replace("OldRange", "NewRange");

                        // Assign the updated formula back to the cell
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Save the modified workbook to a new file (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
