using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data and a formula
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Disable automatic recalculation:
        //   - Set calculation mode to Manual
        //   - Prevent recalculation on save
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Save the workbook without calculating formulas
        workbook.Save("ManualCalc.xlsx");

        // When you need to evaluate formulas, do it explicitly
        workbook.CalculateFormula();

        // Save the workbook after manual calculation
        workbook.Save("ManualCalc_WithResults.xlsx");
    }
}