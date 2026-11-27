// Title: Freeze Excel TODAY() formulas by replacing them with a static date string using Aspose.Cells for .NET (C#)
// Common Searches: C# code to replace TODAY() function with a fixed date in an Excel file using Aspose.Cells | How to freeze dynamic TODAY() formulas in a workbook with Aspose.Cells for .NET | Replace Excel TODAY() formulas with static date strings programmatically in C# | Aspose.Cells example for converting TODAY() formulas to constant values | Iterate worksheets and cells to change TODAY() to a literal date using Aspose.Cells
// Tags: convert TODAY() formula to literal date Aspose.Cells | static date substitution in Excel .NET | modify Excel formulas programmatically C# | iterate workbook cells to change formulas Aspose.Cells | freeze dynamic date function in spreadsheet

using System;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, walks through every worksheet and cell, detects formulas containing the TODAY() function, replaces each formula with the current date formatted as yyyy‑MM‑dd (stored as a constant value), and saves the updated workbook to output.xlsx.
class FreezeTodayFunction
{
    static void Main()
    {
        // Load the workbook (use the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the static date string to replace TODAY()
        string staticDate = DateTime.Today.ToString("yyyy-MM-dd");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                // Check if the cell has a formula and contains the TODAY() function
                if (cell.IsFormula && cell.Formula != null && cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace the formula with the static date string
                    // This freezes the calculation by storing the date as a constant value
                    cell.PutValue(staticDate);
                }
            }
        }

        // Save the modified workbook (use the provided save rule)
        workbook.Save("output.xlsx");
    }
}
