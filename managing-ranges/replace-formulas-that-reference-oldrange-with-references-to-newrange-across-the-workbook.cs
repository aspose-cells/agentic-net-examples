using System;
using Aspose.Cells;

class ReplaceFormulaRange
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all cells in the current worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Process only cells that contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Replace the old named range with the new one inside the formula
                    string updatedFormula = cell.Formula.Replace("OldRange", "NewRange");

                    // Update the cell's formula only if a change was made
                    if (updatedFormula != cell.Formula)
                    {
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}