using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add sample data and formulas
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);               // Simple value
        sheet.Cells["B1"].Formula = "=A1*2";         // Dependent formula
        sheet.Cells["C1"].Formula = "=B1+10";        // Another dependent formula

        // Set the calculation mode to Manual so formulas are not auto‑calculated
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Create calculation options (default settings are fine for this example)
        CalculationOptions calcOptions = new CalculationOptions();

        // Explicitly calculate all formulas using the options
        workbook.CalculateFormula(calcOptions);

        // Display the calculated results
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].IntValue);
        Console.WriteLine("B1 value (A1*2): " + sheet.Cells["B1"].IntValue);
        Console.WriteLine("C1 value (B1+10): " + sheet.Cells["C1"].IntValue);

        // Save the workbook to verify the results are persisted
        workbook.Save("ManualCalculationResult.xlsx");
    }
}