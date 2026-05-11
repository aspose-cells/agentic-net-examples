using System;
using Aspose.Cells;

class SetCalculationModeAfterMerge
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Set a formula that depends on the range that will be merged
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Merge the cells A1:A3 (vertical merge)
        sheet.Cells.Merge(0, 0, 3, 1); // rows 0-2, columns 0-0 (A1:A3)

        // After merging, set calculation mode to Automatic so dependent formulas recalculate instantly
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Force calculation of all formulas now
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MergedAndAutoCalc.xlsx");
    }
}