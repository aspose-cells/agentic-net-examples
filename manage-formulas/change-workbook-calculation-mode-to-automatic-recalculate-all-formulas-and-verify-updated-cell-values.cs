using System;
using Aspose.Cells;

namespace AsposeCellsCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with values and formulas
            cells["A1"].PutValue(5);                 // Base value
            cells["B1"].Formula = "=A1*2";           // Should be 10
            cells["C1"].Formula = "=B1+10";          // Should be 20

            // Set calculation mode to Automatic
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Verify and display the updated cell values
            Console.WriteLine($"A1 (value): {cells["A1"].IntValue}");   // Expected: 5
            Console.WriteLine($"B1 (formula result): {cells["B1"].IntValue}"); // Expected: 10
            Console.WriteLine($"C1 (formula result): {cells["C1"].IntValue}"); // Expected: 20

            // Optionally save the workbook (demonstrates lifecycle usage)
            workbook.Save("CalculationModeAutomatic.xlsx");
        }
    }
}