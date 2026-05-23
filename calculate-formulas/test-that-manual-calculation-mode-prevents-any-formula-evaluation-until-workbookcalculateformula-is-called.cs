using System;
using Aspose.Cells;

namespace ManualCalculationModeTest
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);

            // Set a formula that depends on the above cells
            cells["C1"].Formula = "=A1+B1";

            // Set calculation mode to Manual – formulas will not be evaluated automatically
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // At this point the formula result has not been calculated.
            // The cell's Value will be null (or default) because no calculation has occurred.
            Console.WriteLine("Before CalculateFormula:");
            Console.WriteLine($"C1.Value = {cells["C1"].Value ?? "null"}"); // Expected: null

            // Explicitly calculate all formulas in the workbook
            workbook.CalculateFormula();

            // After calculation the formula result should be available.
            Console.WriteLine("\nAfter CalculateFormula:");
            Console.WriteLine($"C1.Value = {cells["C1"].Value}"); // Expected: 30

            // Optional: Save the workbook to verify that the setting is persisted
            string filePath = "ManualCalcModeDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"\nWorkbook saved to '{filePath}'.");
        }
    }
}