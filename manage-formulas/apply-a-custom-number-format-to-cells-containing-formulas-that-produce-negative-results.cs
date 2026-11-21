// Title: Applying a custom red-negative number format to cells with formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply custom number format to negative formula results | set red negative number style for calculated cells with Aspose.Cells | C# iterate worksheets and format cells that evaluate to negative using Aspose.Cells | how to format negative values in formula cells with Aspose.Cells .NET | apply custom number format after calculating formulas Aspose.Cells C#
// Tags: custom negative number format Aspose.Cells | style formula cells based on calculated value | evaluate formulas before applying cell style | apply red negative formatting to Excel cells .NET | iterate all worksheets and set cell style Aspose

using Aspose.Cells;
using System;

// The example loads an Excel workbook, forces formula calculation, defines a custom number format that shows negative numbers in red, iterates through every worksheet and cell, checks for formula cells with numeric results, applies the custom format only when the result is negative, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are evaluated
        workbook.CalculateFormula();

        // Custom number format: positive numbers normal, negatives in red
        string customNumberFormat = "#,##0.00;[Red]-#,##0.00";

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula)
                {
                    // Check if the calculated result is numeric
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;

                        // Apply custom format only when the result is negative
                        if (value < 0)
                        {
                            Style style = cell.GetStyle();
                            style.Custom = customNumberFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }
        }

        // Save the modified workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
