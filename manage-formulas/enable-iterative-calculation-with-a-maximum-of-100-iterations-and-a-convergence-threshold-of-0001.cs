// Title: Aspose.Cells for .NET: Enable Iterative Calculation (max 100 iterations, 0.001 convergence)
// Description: C# example that activates iterative calculation in Aspose.Cells, sets MaxIteration = 100 and MaxChange = 0.001, creates a circular reference (A1 = A2+1, A2 = A1+1), recalculates the workbook, reads the converged values, and saves the file as IterativeCalculationDemo.xlsx.
// Keywords: Aspose.Cells | iterative calculation | MaxIteration | MaxChange | circular reference | C# .NET | Workbook.CalculateFormula | formula settings | Excel iterative formulas | Aspose.Cells example
// Common Searches: Aspose.Cells enable iterative calculation | set MaxIteration Aspose.Cells .NET | Aspose.Cells convergence threshold | circular reference handling Aspose.Cells | calculate formulas with iteration limit Aspose.Cells
// Developer Intent: Turn on iterative formula evaluation, define iteration and tolerance limits, and compute circular formulas in a .NET workbook.
// Use Cases: Resolve circular references in financial or engineering models. | Run iterative simulations where formulas depend on each other. | Export converged results to Excel for reporting or downstream processing. | Automate batch processing of workbooks that require a fixed iteration cap.
// AI Prompts: Generate C# code that enables iterative calculation in Aspose.Cells with MaxIteration = 100 and MaxChange = 0.001. | Explain how Aspose.Cells handles circular references using iterative calculation and show sample output values. | Show how to retrieve the final values of cells involved in a circular reference after calling Workbook.CalculateFormula().

using System;
using Aspose.Cells;

// C# example that activates iterative calculation in Aspose.Cells, sets MaxIteration = 100 and MaxChange = 0.001, creates a circular reference (A1 = A2+1, A2 = A1+1), recalculates the workbook, reads the converged values, and saves the file as IterativeCalculationDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation with desired settings
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum 100 iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

        // Set up a circular reference to demonstrate iterative calculation
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Display the calculated values
        Console.WriteLine("A1 = " + sheet.Cells["A1"].Value);
        Console.WriteLine("A2 = " + sheet.Cells["A2"].Value);

        // Save the workbook to a file
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
