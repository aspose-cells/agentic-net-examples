using System;
using Aspose.Cells;

class CircularReferenceDetectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set up a circular reference: A1 depends on B1 and B1 depends on A1
        cells["A1"].Formula = "=B1+1";
        cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to detect and resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

        // Disable recursive evaluation to avoid infinite recursion loops
        CalculationOptions calcOptions = new CalculationOptions
        {
            Recursive = false
        };

        // Perform the calculation with the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the calculated values
        Console.WriteLine("A1 = " + cells["A1"].Value);
        Console.WriteLine("B1 = " + cells["B1"].Value);

        // Save the workbook
        workbook.Save("CircularReferenceDemo.xlsx");
    }
}