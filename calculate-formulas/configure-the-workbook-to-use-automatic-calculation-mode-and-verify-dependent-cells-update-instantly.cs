using System;
using Aspose.Cells;

namespace AsposeCellsAutomaticCalcDemo
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

            // Set initial values and a dependent formula
            cells["A1"].PutValue(5);               // Base cell
            cells["A2"].Formula = "=A1*2";         // Dependent cell

            // Configure the workbook to use Automatic calculation mode
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Perform the first calculation
            workbook.CalculateFormula();

            // Verify that the dependent cell reflects the correct result
            Console.WriteLine("After initial calculation:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 (formula result) = {cells["A2"].Value}"); // Expected 10

            // Change the base cell value
            cells["A1"].PutValue(10);

            // Recalculate to see the dependent cell update instantly
            workbook.CalculateFormula();

            // Verify the updated result
            Console.WriteLine("\nAfter updating A1 to 10 and recalculating:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 (formula result) = {cells["A2"].Value}"); // Expected 20

            // Save the workbook (optional)
            workbook.Save("AutomaticCalculationDemo.xlsx");
        }
    }
}