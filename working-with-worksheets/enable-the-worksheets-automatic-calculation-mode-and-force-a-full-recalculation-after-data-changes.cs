using System;
using Aspose.Cells;

namespace AsposeCellsAutomaticCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add initial data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            // Set a formula that sums A1 and A2
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Enable automatic calculation mode
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
            // Force a full recalculation each time a calculation is triggered
            workbook.Settings.FormulaSettings.ForceFullCalculation = true;

            // Perform initial calculation
            workbook.CalculateFormula();

            Console.WriteLine($"Initial A3 value (should be 30): {cells["A3"].Value}");

            // Change data to trigger recalculation
            cells["A1"].PutValue(40);
            cells["A2"].PutValue(60);

            // Recalculate formulas (full recalculation due to ForceFullCalculation)
            workbook.CalculateFormula();

            Console.WriteLine($"After change A3 value (should be 100): {cells["A3"].Value}");

            // Save the workbook (lifecycle rule)
            workbook.Save("AutomaticCalculationDemo.xlsx");
        }
    }
}