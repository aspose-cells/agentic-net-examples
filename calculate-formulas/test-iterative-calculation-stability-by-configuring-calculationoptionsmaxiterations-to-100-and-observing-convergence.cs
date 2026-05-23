using System;
using Aspose.Cells;

class IterativeCalculationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation and set maximum iterations to 100
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;
        // Optional: define the convergence threshold
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Set up a circular reference: A1 depends on A2 and vice‑versa
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        cells["A1"].Formula = "=A2+1";
        cells["A2"].Formula = "=A1+1";

        // Perform the calculation
        workbook.CalculateFormula();

        // Display the resulting values (should converge after the specified iterations)
        Console.WriteLine("A1 value: " + cells["A1"].Value);
        Console.WriteLine("A2 value: " + cells["A2"].Value);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}