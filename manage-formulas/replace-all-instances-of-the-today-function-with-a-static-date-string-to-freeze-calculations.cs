// Title: Freeze TODAY() formulas to a static date in Excel with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans every worksheet and used cell, detects formulas containing the TODAY() function, replaces each with the current date (DateTime.Today) as a constant value, and saves the result to a new file.
// Keywords: Aspose.Cells C# | replace TODAY formula | static date Excel | freeze dynamic formulas | convert TODAY() to constant | Excel date freeze .NET | programmatic Excel date replacement
// Common Searches: how to replace TODAY() with a fixed date using Aspose.Cells | freeze TODAY function in Excel programmatically C# | Aspose.Cells replace dynamic date formulas | convert TODAY() to static value in .NET | remove volatile TODAY() formulas from workbook
// Developer Intent: Replace all TODAY() formulas in a workbook with a constant date value.
// Use Cases: Create a financial snapshot that retains the generation date after distribution. | Archive Excel reports for compliance where dates must not change on opening. | Prepare regulatory submissions with all formulas frozen to static values.
// AI Prompts: Show how to replace TODAY() with a custom date string instead of DateTime.Today. | Demonstrate using Aspose.Cells' ReplaceFormula method to freeze TODAY() across a workbook. | Explain handling of array formulas or named ranges that reference TODAY() during replacement.

using System;
using Aspose.Cells;

// Loads an Excel workbook, scans every worksheet and used cell, detects formulas containing the TODAY() function, replaces each with the current date (DateTime.Today) as a constant value, and saves the result to a new file.
class FreezeTodayFunction
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the static date to replace TODAY() with (today's date at runtime)
        DateTime staticDate = DateTime.Today;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            Cells cells = worksheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a formula that uses TODAY()
                    if (cell.IsFormula && cell.Formula != null &&
                        cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the formula with the static date value and remove the formula
                        cell.PutValue(staticDate);
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
