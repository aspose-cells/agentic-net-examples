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
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Enable iterative calculation and set custom thresholds
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true; // turn on iterative mode
            workbook.Settings.FormulaSettings.MaxIteration = 200;               // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.0005;               // custom maximum change for convergence

            // Create a circular reference for demonstration
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Perform calculation using the configured settings
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].Value);
            Console.WriteLine("Maximum change used: " + workbook.Settings.FormulaSettings.MaxChange);

            // Save the workbook (optional)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}