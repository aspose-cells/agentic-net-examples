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

            // Add sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that sums the three values
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Ensure calculation mode is Automatic before merging
            // (FormulaSettings.CalculationMode uses the CalcModeType enum)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform a merge that includes the formula cell (merge rule)
            // Merge cells B1:C1 – the formula stays in the upper‑left cell (B1)
            cells.Merge(0, 1, 1, 2);

            // Keep calculation mode Automatic after merging (re‑apply if needed)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Save the workbook (save rule)
            workbook.Save("FormulaMergeDemo.xlsx");
        }
    }
}