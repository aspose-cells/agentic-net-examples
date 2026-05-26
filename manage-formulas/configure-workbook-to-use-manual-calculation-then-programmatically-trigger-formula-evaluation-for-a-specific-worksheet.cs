using Aspose.Cells;
using System;

class ManualCalculationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data and a formula
        worksheet.Cells["A1"].PutValue(5);
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B1"].Formula = "=A1+A2";

        // Set the calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Create calculation options (default settings)
        CalculationOptions calcOptions = new CalculationOptions();

        // Trigger calculation for this specific worksheet (recursive = true)
        worksheet.CalculateFormula(calcOptions, true);

        // Display the calculated result
        Console.WriteLine("B1 value after manual calculation: " + worksheet.Cells["B1"].Value);

        // Save the workbook (optional)
        workbook.Save("ManualCalculationDemo.xlsx");
    }
}