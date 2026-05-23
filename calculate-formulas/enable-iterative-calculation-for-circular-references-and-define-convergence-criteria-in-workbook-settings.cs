using System;
using Aspose.Cells;

namespace IterativeCalculationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set up a circular reference: A1 depends on B1 and B1 depends on A1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to resolve the circular reference
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Define convergence criteria
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // acceptable change between iterations

            // Perform the calculation
            workbook.CalculateFormula();

            // Output the results after calculation
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].DoubleValue);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].DoubleValue);

            // Save the workbook (you can change the path as needed)
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}