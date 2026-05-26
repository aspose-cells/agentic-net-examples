using System;
using Aspose.Cells;

class SetPrecisionAsDisplayedDemo
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

        // Set number format for A1 to show 2 decimal places (0.00)
        Style style = cells["A1"].GetStyle();
        style.Number = 2;
        cells["A1"].SetStyle(style);

        // Set a formula in B1 that references A1
        cells["B1"].Formula = "=A1";

        // Calculate formulas; B1 will use the rounded value because PrecisionAsDisplayed is true
        workbook.CalculateFormula();

        // Output the displayed values
        Console.WriteLine("A1 displayed value: " + cells["A1"].StringValue); // Expected: 1.23
        Console.WriteLine("B1 calculated value: " + cells["B1"].Value);      // Expected: 1.23

        // Save the workbook
        workbook.Save("PrecisionAsDisplayedDemo.xlsx");
    }
}