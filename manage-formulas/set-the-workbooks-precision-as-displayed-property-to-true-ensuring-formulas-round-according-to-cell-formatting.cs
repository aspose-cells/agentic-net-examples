// Title: Aspose.Cells for .NET – Enable PrecisionAsDisplayed to round formulas to cell display format (C#)
// Description: This C# example creates a workbook, activates the PrecisionAsDisplayed setting, formats cell A1 to two decimal places, assigns a formula in B1 that references A1, recalculates the workbook, and demonstrates that both the displayed value of A1 and the calculated result in B1 are rounded to the displayed precision before saving the file.
// Keywords: Aspose.Cells | PrecisionAsDisplayed | C# | .NET | formula rounding | displayed precision | cell formatting rounding | workbook.Settings.FormulaSettings | financial spreadsheet rounding
// Common Searches: Aspose.Cells PrecisionAsDisplayed C# | How to round formula results to displayed decimals in Aspose.Cells | Enable displayed precision in .NET Excel library | C# calculate workbook with displayed precision | Aspose.Cells rounding based on cell format
// Developer Intent: Activate workbook.Settings.FormulaSettings.PrecisionAsDisplayed so that all formula calculations use the displayed cell precision, ensuring results are rounded according to the cell's number format.
// Use Cases: Financial reports that require values limited to two decimal places | Generating Excel invoices where dependent cells must share the same rounding | Preparing data for downstream systems that accept only rounded numbers | Creating spreadsheet templates where visual precision must match stored values
// AI Prompts: Show C# code using Aspose.Cells to set PrecisionAsDisplayed true, format a cell to two decimals, add a formula referencing that cell, and display the rounded result. | Provide an Aspose.Cells .NET example that enables displayed precision, applies number formatting, and verifies formula rounding in the generated workbook. | Explain the impact of PrecisionAsDisplayed on formula calculation and how to test it with a simple C# workbook.

using System;
using Aspose.Cells;

// This C# example creates a workbook, activates the PrecisionAsDisplayed setting, formats cell A1 to two decimal places, assigns a formula in B1 that references A1, recalculates the workbook, and demonstrates that both the displayed value of A1 and the calculated result in B1 are rounded to the displayed precision before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable PrecisionAsDisplayed so calculations use the displayed precision
        workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Put a value with many decimal places into A1
        cells["A1"].PutValue(1.23456);

        // Set the display format of A1 to show 2 decimal places (0.00)
        Style style = cells["A1"].GetStyle();
        style.Number = 2;
        cells["A1"].SetStyle(style);

        // Set a formula in B1 that references A1
        cells["B1"].Formula = "=A1";

        // Calculate formulas; the result will respect the displayed precision
        workbook.CalculateFormula();

        // Output the displayed value of A1 and the calculated value of B1
        Console.WriteLine("A1 displayed value: " + cells["A1"].StringValue); // Expected: 1.23
        Console.WriteLine("B1 calculated value: " + cells["B1"].Value);      // Expected: 1.23

        // Save the workbook (optional)
        workbook.Save("PrecisionAsDisplayedDemo.xlsx");
    }
}
