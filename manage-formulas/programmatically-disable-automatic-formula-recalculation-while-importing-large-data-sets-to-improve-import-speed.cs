using System;
using Aspose.Cells;

class DisableAutoCalcDuringImport
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Disable automatic formula calculation
        wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
        wb.Settings.FormulaSettings.CalculateOnOpen = false;
        wb.Settings.FormulaSettings.CalculateOnSave = false;

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Simulate importing a large data set
        int rows = 10000;
        int cols = 10;
        object[] data = new object[rows * cols];
        for (int i = 0; i < data.Length; i++)
        {
            data[i] = i; // sample numeric data
        }

        // Import data starting at cell A1 (row 0, column 0)
        // false = horizontal import (fill rows first)
        cells.ImportObjectArray(data, 0, 0, false);

        // After import, calculate formulas explicitly if required
        wb.CalculateFormula();

        // Save the workbook
        wb.Save("LargeDataManualCalc.xlsx", SaveFormat.Xlsx);
    }
}