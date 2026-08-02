using System;
using Aspose.Cells;

namespace ManualCalcModeTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put some values
            cells["A1"].PutValue(5);
            cells["B1"].PutValue(10);

            // Set a formula that depends on the above values
            cells["C1"].Formula = "=A1+B1";

            // Set calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // At this point, the formula should NOT be evaluated automatically
            Console.WriteLine("Before manual calculation:");
            Console.WriteLine($"C1 value (expected empty or 0): {cells["C1"].Value ?? "null"}");

            // Perform manual calculation
            workbook.CalculateFormula();

            // After calculation, the formula result should be available
            Console.WriteLine("After manual calculation:");
            Console.WriteLine($"C1 value (expected 15): {cells["C1"].Value}");

            // Save the workbook (optional, demonstrates that saving does not trigger calculation)
            workbook.Save("ManualCalcModeTest.xlsx");
        }
    }
}