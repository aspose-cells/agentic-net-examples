// Title: Apply a Red Negative Number Format to Formula Cells with Aspose.Cells for .NET
// Description: Creates a workbook, inserts numeric values and formulas, calculates them, defines a custom format that shows negatives in red, and uses a StyleFlag to apply the format only to formula cells whose evaluated value is negative before saving the file.
// Keywords: Aspose.Cells | custom number format | negative numbers | formula results | StyleFlag | C# | .NET | workbook styling | apply style to cells
// Common Searches: Aspose.Cells format negative formula results | apply custom number format to negative values C# | StyleFlag number format example Aspose.Cells | highlight negative numbers in Excel using Aspose.Cells | C# code to style cells with negative formula outcomes
// Developer Intent: Apply a custom number format that highlights only the negative results of formulas while leaving other cells unchanged.
// Use Cases: Financial reports where losses are shown in red for quick visual identification. | Audit spreadsheets that automatically flag negative calculated values. | Performance dashboards that emphasize metrics falling below zero without using conditional formatting.
// AI Prompts: Show how to extend the sample to format zero values in gray. | Provide a solution that uses Aspose.Cells Conditional Formatting to highlight negative results. | Explain how to apply the same custom format to an entire column with a single API call.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, inserts numeric values and formulas, calculates them, defines a custom format that shows negatives in red, and uses a StyleFlag to apply the format only to formula cells whose evaluated value is negative before saving the file.
class ApplyCustomNumberFormatToNegativeFormulaResults
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some cells with values and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=A1-A2";   // Result: -10
            cells["B2"].Formula = "=A2-A1";   // Result: 10
            cells["C1"].Formula = "=A1*-1";   // Result: -10
            cells["C2"].Formula = "=A2*2";    // Result: 40

            // Calculate all formulas so that cell values are up‑to‑date
            workbook.CalculateFormula();

            // Define a custom number format that shows negatives in red
            string customNumberFormat = "_-\"$\"* #,##0.00;[Red]-\"$\"* #,##0.00;_-\"$\"* \"-\"??_;_@_";

            // Create a style that contains only the custom number format
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = customNumberFormat;

            // Use a StyleFlag to apply only the number format part of the style
            StyleFlag numberFormatFlag = new StyleFlag();
            numberFormatFlag.NumberFormat = true;

            // Iterate through all cells, find those with formulas that evaluate to a negative number,
            // and apply the custom number format using the flag
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && cell.Type == CellValueType.IsNumeric && cell.DoubleValue < 0)
                {
                    // Apply the style to the single cell range
                    AsposeRange range = cells.CreateRange(cell.Row, cell.Column, 1, 1);
                    range.ApplyStyle(customStyle, numberFormatFlag);
                }
            }

            // Save the workbook
            workbook.Save("NegativeFormulaNumberFormat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
