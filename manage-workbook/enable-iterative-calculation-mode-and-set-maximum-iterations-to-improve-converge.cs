using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation to resolve circular references
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Set the maximum number of iterations for convergence
            workbook.Settings.FormulaSettings.MaxIteration = 100;

            // (Optional) Define the maximum change threshold for convergence
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a circular reference scenario
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Perform formula calculation with the iterative settings
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + cells["A2"].Value);

            // Save the workbook (adjust the path as needed)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}