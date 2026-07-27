using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");

        // Disable automatic calculation for tables only.
        // Other formulas will still be calculated automatically.
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

        // Save the workbook with the new setting.
        workbook.Save("output.xlsx");
    }
}