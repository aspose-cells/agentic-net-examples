// Title: Aspose.Cells .NET – Enable Iterative Calculation to Resolve Circular References
// Description: Shows how to activate iterative calculation in Aspose.Cells for .NET, configure MaxIteration and MaxChange, create a circular reference between cells A1 and B1, run workbook.CalculateFormula(), and save the workbook. This lets formulas with circular dependencies converge safely.
// Keywords: Aspose.Cells iterative calculation | circular reference handling | FormulaSettings.EnableIterativeCalculation | MaxIteration Aspose.Cells | MaxChange Aspose.Cells | .NET spreadsheet calculation | Aspose.Cells CalculateFormula | iterative formula evaluation
// Common Searches: how to enable iterative calculation in Aspose.Cells | Aspose.Cells circular reference example | set MaxIteration and MaxChange in Aspose.Cells | calculate formulas with circular dependencies using Aspose.Cells | iterative calculation settings .NET
// Developer Intent: Turn on iterative calculation and define iteration limits to evaluate circular formulas safely.
// Use Cases: Process financial models that contain circular references by enabling iterative calculation before running calculations. | Prevent endless loops in large workbooks by setting MaxIteration and MaxChange, then invoking CalculateFormula to achieve convergence. | Generate a spreadsheet with self‑referencing formulas, compute values automatically, and export the result for downstream reporting.
// AI Prompts: Provide C# code that enables iterative calculation in Aspose.Cells, sets MaxIteration to 200 and MaxChange to 0.0001, and prints the values of circular cells. | Explain how MaxIteration and MaxChange affect convergence when using Aspose.Cells iterative calculation. | Show an example of handling a circular reference (A1 = B1+1, B1 = A1+1) with Aspose.Cells and saving the workbook.

using System;
using Aspose.Cells;

// Shows how to activate iterative calculation in Aspose.Cells for .NET, configure MaxIteration and MaxChange, create a circular reference between cells A1 and B1, run workbook.CalculateFormula(), and save the workbook. This lets formulas with circular dependencies converge safely.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation to resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Optional: configure iteration limits
        workbook.Settings.FormulaSettings.MaxIteration = 100;
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Set up a circular reference for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Perform calculation
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
        Console.WriteLine("B1 value: " + sheet.Cells["B1"].Value);

        // Save the workbook (optional)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
