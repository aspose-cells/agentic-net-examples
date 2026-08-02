using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty)
        Workbook workbook = new Workbook();

        // Set the calculation mode to Automatic (for Excel compatibility)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Ensure a full recalculation is performed each time
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;

        // Create default calculation options
        CalculationOptions calcOptions = new CalculationOptions();

        // Trigger calculation of all formulas in the workbook
        workbook.CalculateFormula(calcOptions);

        // Save the workbook (optional, demonstrates that formulas are now calculated)
        workbook.Save("CalculatedWorkbook.xlsx");
    }
}