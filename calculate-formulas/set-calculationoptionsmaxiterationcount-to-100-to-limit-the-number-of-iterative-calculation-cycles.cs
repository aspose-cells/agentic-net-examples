using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation and limit it to 100 iterations
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Sample circular reference to illustrate the effect
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Perform calculation using the configured settings
        workbook.CalculateFormula();

        // Display the results
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
        Console.WriteLine("A2 value: " + sheet.Cells["A2"].Value);

        // Save the workbook
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}