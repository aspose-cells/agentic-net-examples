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

            // Set the maximum number of iterations (100)
            workbook.Settings.FormulaSettings.MaxIteration = 100;

            // Optional: define the maximum change threshold for convergence
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a circular reference: A1 depends on A2, A2 depends on A1
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated values after iterative processing
            Console.WriteLine("A1 value: " + cells["A1"].Value);
            Console.WriteLine("A2 value: " + cells["A2"].Value);

            // Save the workbook to verify the results (optional)
            workbook.Save("IterativeCircularDemo.xlsx");
        }
    }
}