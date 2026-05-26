using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Enable iterative calculation to resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Set the maximum number of iterations to 100
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Set the convergence threshold (maximum change) to 0.001
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Example circular reference to demonstrate iterative calculation
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
        Console.WriteLine("A2 value: " + sheet.Cells["A2"].Value);

        // Save the workbook (lifecycle: save)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}