// Title: C# – Enable Iterative Calculation for Circular References with Aspose.Cells
// Description: Demonstrates how to activate iterative calculation in Aspose.Cells for .NET, configure MaxIteration and MaxChange, evaluate circular formulas, retrieve results, and save the workbook.
// Keywords: Aspose.Cells iterative calculation | circular reference formula .NET | EnableIterativeCalculation C# | MaxIteration Aspose.Cells | MaxChange Aspose.Cells | Aspose.Cells formula settings | calculate workbook with recursion
// Common Searches: how to enable iterative calculation in Aspose.Cells C# | Aspose.Cells circular reference handling | set MaxIteration and MaxChange Aspose.Cells | iterative formula evaluation Aspose.Cells .NET | resolve circular formulas with Aspose.Cells
// Developer Intent: Turn on iterative calculation so formulas that reference each other can be computed without errors.
// Use Cases: Financial models that contain circular dependencies, such as interest accrual loops. | Engineering spreadsheets requiring convergence of iterative formulas. | Automated reports that must resolve recursive calculations before export.
// AI Prompts: Generate C# code using Aspose.Cells to enable iterative calculation with custom MaxIteration and MaxChange values for a workbook containing circular formulas. | Explain the impact of MaxIteration and MaxChange settings on convergence speed and accuracy in Aspose.Cells iterative calculations. | Show how to retrieve the number of iterations performed after calling CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to activate iterative calculation in Aspose.Cells for .NET, configure MaxIteration and MaxChange, evaluate circular formulas, retrieve results, and save the workbook.
class EnableIterativeCalculationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up a circular reference for demonstration
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        // Optional: define iteration limits and tolerance
        workbook.Settings.FormulaSettings.MaxIteration = 100;
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Perform calculation
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value after calculation: " + sheet.Cells["A1"].Value);
        Console.WriteLine("B1 value after calculation: " + sheet.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
