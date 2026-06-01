using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set up a circular reference for demonstration
        worksheet.Cells["A1"].Formula = "=B1+1";
        worksheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Optional: configure iteration limits and tolerance
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // maximum change per iteration

        // Perform the calculation
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value after calculation: " + worksheet.Cells["A1"].Value);
        Console.WriteLine("B1 value after calculation: " + worksheet.Cells["B1"].Value);

        // Save the workbook (demonstrates usage of the save rule)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}