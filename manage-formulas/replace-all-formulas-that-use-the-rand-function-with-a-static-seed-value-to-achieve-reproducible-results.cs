using System;
using Aspose.Cells;

class ReplaceRandWithStatic
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the static value that will replace RAND()
        const string staticValue = "0.12345";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Loop through each cell that contains a formula
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Formula != null && cell.Formula.IndexOf("RAND()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace all occurrences of RAND() with the static value
                    string newFormula = cell.Formula.Replace("RAND()", staticValue, StringComparison.OrdinalIgnoreCase);

                    // Set the updated formula back to the cell
                    cell.Formula = newFormula;
                }
            }
        }

        // Recalculate all formulas after the replacements
        workbook.CalculateFormula();

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}