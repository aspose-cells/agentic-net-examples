// Title: Apply a custom red number format to formula cells that evaluate to negative values using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# format cells with negative formula results | set custom number format for negative values in calculated cells using Aspose.Cells | how to apply red color to negative numbers produced by formulas in Aspose.Cells | iterate over cells with formulas and style negative results in .NET | apply conditional number format to formula cells after calculation Aspose.Cells
// Tags: custom negative number format Aspose.Cells | style formula cells based on result Aspose.Cells | apply red negative format Excel .NET | iterate cells with formulas Aspose.Cells | conditional number formatting after calculation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds values and formulas, calculates all formulas, then loops through each cell. If a cell contains a formula whose evaluated result is negative, it assigns the custom number format "#,##0.00;[Red]-#,##0.00" to display negative numbers in red, and finally saves the file as FormattedNegativeResults.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data and formulas
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "A1-A2"; // Result will be -10
        cells["B2"].Formula = "A2-A1"; // Result will be 10

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Custom number format: positive numbers normal, negative numbers in red
        string customFormat = "#,##0.00;[Red]-#,##0.00";

        // Apply the custom format to cells with formulas that evaluate to a negative value
        foreach (Cell cell in cells)
        {
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                // Try to parse the calculated value as a double
                if (double.TryParse(cell.Value?.ToString(), out double result) && result < 0)
                {
                    Style style = cell.GetStyle();
                    style.Custom = customFormat;
                    cell.SetStyle(style);
                }
            }
        }

        // Save the workbook
        workbook.Save("FormattedNegativeResults.xlsx");
    }
}
