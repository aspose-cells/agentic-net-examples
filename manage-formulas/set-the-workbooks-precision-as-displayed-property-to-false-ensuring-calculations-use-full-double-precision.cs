using Aspose.Cells;
using System;

class SetPrecisionAsDisplayedDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable PrecisionAsDisplayed to use full double precision in calculations
        workbook.Settings.FormulaSettings.PrecisionAsDisplayed = false;

        // Add sample data and a formula to demonstrate the effect
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(1.23456789);
        sheet.Cells["A2"].PutValue(2.34567891);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Perform calculation with full precision
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("PrecisionAsDisplayedFalse.xlsx");
    }
}