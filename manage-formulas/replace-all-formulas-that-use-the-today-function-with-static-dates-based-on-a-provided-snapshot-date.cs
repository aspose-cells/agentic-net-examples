// Title: Replace TODAY() formulas with a static snapshot date in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file with Aspose.Cells, locate cells whose formula contains TODAY(), and set those cells to a provided DateTime snapshot. | Traverse all worksheets and cells in a workbook, replace any TODAY() function in formulas with a fixed date value, then save the updated file. | Programmatically convert dynamic TODAY() formulas to constant dates using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# replace TODAY() function with a specific date | How to convert dynamic TODAY() formulas to static dates in an .xlsx using .NET | C# code example for substituting TODAY() formulas with a snapshot date in Excel | Replace Excel TODAY() formulas with constant date programmatically Aspose.Cells | Iterate through worksheets and replace TODAY() formulas with fixed date in C#
// Tags: replace TODAY() formula Aspose.Cells | static snapshot date Excel C# | Aspose.Cells formula manipulation | iterate worksheets Aspose.Cells | convert dynamic date function to constant .xlsx

using System;
using Aspose.Cells;

// Loads an Excel workbook, scans every cell for formulas containing TODAY(), replaces those formulas with a predefined snapshot DateTime, and saves the modified file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Define the snapshot date that will replace TODAY()
        DateTime snapshotDate = new DateTime(2023, 10, 1); // adjust as needed

        // Load the workbook (use the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all used cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell contains a formula that uses TODAY()
                if (cell.IsFormula && cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Replace the formula with the static snapshot date value
                    cell.PutValue(snapshotDate);
                }
            }
        }

        // Save the modified workbook (use the provided save rule)
        workbook.Save("output.xlsx");
    }
}
