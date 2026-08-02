using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the calculation mode to Automatic for immediate formula updates
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Add sample data and a formula to demonstrate automatic calculation
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Calculate formulas now that the mode is Automatic
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("AutomaticCalculation.xlsx");
    }
}