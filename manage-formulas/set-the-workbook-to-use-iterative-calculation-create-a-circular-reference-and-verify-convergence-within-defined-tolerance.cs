using System;
using Aspose.Cells;

namespace IterativeCalculationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a circular reference: A1 depends on B1, B1 depends on A1
            cells["A1"].Formula = "=B1*0.5";
            cells["B1"].Formula = "=A1*0.5";

            // Enable iterative calculation and define convergence settings
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;      // maximum iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;      // tolerance for convergence

            // Perform formula calculation
            workbook.CalculateFormula();

            // Retrieve calculated values
            double a1Value = cells["A1"].DoubleValue;
            double b1Value = cells["B1"].DoubleValue;

            // Verify convergence: the absolute difference should be within the defined tolerance
            double difference = Math.Abs(a1Value - b1Value);
            bool isConverged = difference <= workbook.Settings.FormulaSettings.MaxChange;

            // Output results
            Console.WriteLine($"A1 = {a1Value}");
            Console.WriteLine($"B1 = {b1Value}");
            Console.WriteLine($"Difference = {difference}");
            Console.WriteLine($"Converged within tolerance: {isConverged}");

            // Save the workbook (optional)
            workbook.Save("IterativeCircular.xlsx");
        }
    }
}