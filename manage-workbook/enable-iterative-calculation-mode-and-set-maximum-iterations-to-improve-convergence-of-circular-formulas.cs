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
            workbook.Settings.FormulaSettings.MaxIteration = 20;

            // (Optional) Define the maximum change allowed between iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Set up a simple circular reference: A1 depends on A2 and vice‑versa
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Perform the calculation using the workbook's engine
            workbook.CalculateFormula();

            // Display the calculated values
            Console.WriteLine("A1 value: " + cells["A1"].Value);
            Console.WriteLine("A2 value: " + cells["A2"].Value);

            // Save the workbook (optional)
            workbook.Save("IterativeDemo.xlsx");
        }
    }
}