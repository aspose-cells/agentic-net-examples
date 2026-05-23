using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the calculation mode to Automatic
        // This setting is saved in the file so that Excel will recalculate automatically when opened
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Add sample data and a formula to demonstrate the setting
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Optionally calculate formulas now (Aspose.Cells does not auto‑calculate)
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("AutomaticCalculation.xlsx");
    }
}