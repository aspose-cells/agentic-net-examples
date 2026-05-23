using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation to handle circular references
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Set the maximum number of iterations (precision tolerance)
            workbook.Settings.FormulaSettings.MaxIteration = 50;   // example: 50 iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.0001; // example: tolerance of 0.0001

            // Create a simple circular reference for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2+1";
            sheet.Cells["A2"].Formula = "=A1+1";

            // Perform calculation with the iterative settings
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("A1 value after iterative calculation: " + sheet.Cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + sheet.Cells["A2"].Value);

            // Save the workbook
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}