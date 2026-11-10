// Title: How to replace Excel TODAY() formulas with a fixed snapshot date using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace TODAY() function with a specific date in an existing workbook | Convert dynamic TODAY() formulas to static dates when processing Excel files with Aspose.Cells | C# code to scan all cells and substitute TODAY() formulas with a snapshot date using Aspose.Cells | Set a fixed date for TODAY() formulas in Excel before saving with Aspose.Cells .NET | Replace volatile date formulas with constant values in Excel via Aspose.Cells API
// Tags: replace TODAY() formula Aspose.Cells | static date substitution Excel C# | iterate worksheet cells Aspose.Cells | clear cell formula after setting value Aspose | snapshot date replacement Excel workbook

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through every used cell, replaces any formula containing TODAY() with a provided snapshot DateTime, clears the formula to make the cell a constant, and saves the updated file.
class Program
{
    static void Main()
    {
        // Snapshot date to replace TODAY() with
        DateTime snapshotDate = new DateTime(2023, 10, 1); // <-- set your snapshot date here

        // Load the workbook (lifecycle rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains a formula with TODAY()
                    if (cell.IsFormula && cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the formula with the static snapshot date
                        cell.PutValue(snapshotDate);
                        // Clear the formula to make it a constant value
                        cell.Formula = string.Empty;
                    }
                }
            }
        }

        // Save the modified workbook (lifecycle rule)
        workbook.Save("output.xlsx");
    }
}
