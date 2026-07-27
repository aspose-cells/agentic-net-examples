using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Switch calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Perform bulk updates (example: fill 1000 rows)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        for (int i = 0; i < 1000; i++)
        {
            // Put a value in column A
            cells[i, 0].PutValue(i + 1);
            // Set a formula in column B that depends on column A
            cells[i, 1].Formula = $"=A{i + 1}*2";
        }

        // Calculate all formulas once after bulk updates
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("BulkUpdateManualCalc.xlsx");
    }
}