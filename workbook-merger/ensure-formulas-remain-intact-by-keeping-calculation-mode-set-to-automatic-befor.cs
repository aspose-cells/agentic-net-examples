using System;
using Aspose.Cells;

namespace AsposeCellsFormulaMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data and a formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=SUM(A1:A2)";

            // Assume the workbook might be in Manual calculation mode.
            // Store the original calculation mode.
            CalcModeType originalMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Set calculation mode to Automatic before merging to keep formulas intact.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform a merge operation (e.g., merge A1:B1)
            cells.Merge(0, 0, 1, 2);

            // Restore the original calculation mode after merging.
            workbook.Settings.FormulaSettings.CalculationMode = originalMode;

            // Save the workbook (using default save options)
            workbook.Save("MergedWithFormulas.xlsx");
        }
    }
}