// Title: How to set maximum iterations and precision tolerance for iterative formula calculation to resolve circular references with Aspose.Cells for .NET
// AI Prompts: Enable iterative calculation, set Workbook.Settings.FormulaSettings.MaxIteration to 50 and MaxChange to 0.0001, then call workbook.CalculateFormula() to resolve circular references. | Create a circular reference between cells A1 and A2, configure iterative settings, calculate the workbook, and save the file using Aspose.Cells in C#.
// Common Searches: Aspose.Cells set MaxIteration for iterative calculations in C# | configure MaxChange tolerance for formula loops using Aspose.Cells .NET | example of enabling iterative formula calculation with Aspose.Cells workbook | how to resolve Excel circular formulas with Aspose.Cells iterative settings | C# code to limit iterative calculation loops in Aspose.Cells
// Tags: maxiteration setting Aspose.Cells | maxchange tolerance Aspose.Cells | circular formula handling Aspose.Cells | enable iterative formulas C# | Aspose.Cells workbook calculation options

using System;
using Aspose.Cells;

// The example creates a new workbook, enables iterative calculation, sets MaxIteration to 50 and MaxChange to 0.0001, defines a circular reference between A1 and A2, runs the calculation, prints the resulting values, and saves the workbook as IterativeCalculationDemo.xlsx.
class IterativeCalculationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation to resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Set the maximum number of iterations
        workbook.Settings.FormulaSettings.MaxIteration = 50;

        // Set the precision tolerance (maximum change)
        workbook.Settings.FormulaSettings.MaxChange = 0.0001;

        // Create a circular reference for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Perform calculation with the configured iterative settings
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 = " + sheet.Cells["A1"].Value);
        Console.WriteLine("A2 = " + sheet.Cells["A2"].Value);

        // Save the workbook
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
