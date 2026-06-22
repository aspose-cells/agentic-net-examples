using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);

        // Add a formula that sums A1 and A2
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Set calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prevent automatic calculation on save
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // At this point the formula has not been evaluated
        Console.WriteLine("Before manual calculation, A3 value: " + sheet.Cells["A3"].Value);

        // Trigger calculation only when needed
        workbook.CalculateFormula();

        // Now the formula result is available
        Console.WriteLine("After manual calculation, A3 value: " + sheet.Cells["A3"].Value);

        // Save the workbook
        workbook.Save("ManualCalcDemo.xlsx");
    }
}