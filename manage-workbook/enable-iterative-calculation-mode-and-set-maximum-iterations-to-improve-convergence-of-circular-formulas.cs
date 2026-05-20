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

            // Set up a simple circular reference scenario
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2+1";
            sheet.Cells["A2"].Formula = "=A1+1";

            // Perform formula calculation using the configured settings
            workbook.CalculateFormula();

            // Display the calculated values
            Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
            Console.WriteLine("A2 value: " + sheet.Cells["A2"].Value);

            // Save the workbook (optional)
            workbook.Save("IterativeDemo.xlsx");
        }
    }
}