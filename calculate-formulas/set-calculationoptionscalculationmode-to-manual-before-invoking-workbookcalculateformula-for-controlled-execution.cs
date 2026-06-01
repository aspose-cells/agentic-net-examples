using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);

        // Add a formula that depends on the data above
        sheet.Cells["B1"].Formula = "=A1+A2";

        // ------------------------------------------------------------
        // Set the calculation mode to Manual so that formulas are NOT
        // calculated automatically. This setting is stored in the
        // FormulaSettings of the workbook.
        // ------------------------------------------------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Create a CalculationOptions instance (optional, can be customized)
        CalculationOptions calcOptions = new CalculationOptions();

        // Explicitly trigger formula calculation now that the mode is Manual
        workbook.CalculateFormula(calcOptions);

        // Display the calculated result
        Console.WriteLine("Result of B1: " + sheet.Cells["B1"].IntValue);

        // Save the workbook (the calculation mode setting will be persisted)
        workbook.Save("ManualCalculationDemo.xlsx");
    }
}