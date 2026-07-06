using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set up a circular reference for demonstration
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to allow circular reference evaluation
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        // Optional: define iteration limits
        workbook.Settings.FormulaSettings.MaxIteration = 100;
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
        Console.WriteLine("B1 value: " + sheet.Cells["B1"].Value);
    }
}