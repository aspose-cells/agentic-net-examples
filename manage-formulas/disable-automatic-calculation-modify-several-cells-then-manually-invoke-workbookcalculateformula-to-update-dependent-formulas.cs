using System;
using Aspose.Cells;

namespace AsposeCellsManualCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some initial data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);
            cells["B1"].PutValue(2);
            cells["B2"].PutValue(3);

            // Set formulas that depend on the above cells
            cells["C1"].Formula = "=A1+B1";   // =5+2 = 7
            cells["C2"].Formula = "=A2*B2";   // =10*3 = 30
            cells["D1"].Formula = "=C1+C2";   // =7+30 = 37

            // Disable automatic calculation (feature: FormulaSettings.CalculationMode)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Modify several cells after disabling auto‑calc
            cells["A1"].PutValue(20);   // Change from 5 to 20
            cells["B2"].PutValue(4);    // Change from 3 to 4

            // Manually trigger formula calculation (feature: Workbook.CalculateFormula)
            workbook.CalculateFormula();

            // Output results to console for verification
            Console.WriteLine("C1 (A1+B1) = " + cells["C1"].Value);
            Console.WriteLine("C2 (A2*B2) = " + cells["C2"].Value);
            Console.WriteLine("D1 (C1+C2) = " + cells["D1"].Value);

            // Save the workbook (lifecycle: save)
            workbook.Save("ManualCalculationResult.xlsx");
        }
    }
}