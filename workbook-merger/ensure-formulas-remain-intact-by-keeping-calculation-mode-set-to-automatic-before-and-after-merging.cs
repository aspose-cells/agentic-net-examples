using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data and a formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=SUM(A1:A2)"; // Formula that will be affected by merging

            // Ensure calculation mode is Automatic before merging
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform a merge operation (merge rule)
            // Merge cells B1:C1 (row 0, column 1) spanning 1 row and 2 columns
            cells.Merge(0, 1, 1, 2);

            // Verify that the calculation mode is still Automatic after merging
            // (re‑apply if needed to guarantee the setting)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Save the workbook (save rule)
            workbook.Save("FormulaMergeResult.xlsx");
        }
    }
}